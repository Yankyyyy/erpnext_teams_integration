let teams_settings_cache = null;

async function get_teams_settings() {
    if (!teams_settings_cache) {
        teams_settings_cache = await frappe.db.get_doc(
            "Teams Settings",
            "Teams Settings"
        );
    }
    return teams_settings_cache;
}

frappe.ui.form.on("Project", {
    async refresh(frm) {
        const settings_doc = await get_teams_settings();

        const enabledSet = new Set(
            (settings_doc.enabled_doctypes || [])
                .map(row => row.doctype_name)
        );

        if (!enabledSet.has(frm.doctype)) {
            frm.remove_custom_button(__('Create Teams Chat'), __("Teams"));
            frm.remove_custom_button(__('Open Teams Chat'), __("Teams"));
            frm.remove_custom_button(__('Send Teams Message'), __("Teams"));
            frm.remove_custom_button(__('Post to Channel'), __("Teams"));
            frm.remove_custom_button(__('Create Teams Meeting'), __("Teams"));
            frm.remove_custom_button(__('Sync Now'), __("Teams"));
            return;
        }

        // 1. Handle custom meeting button override
        if (frm.doc.custom_teams_meeting_url) {
            frm.fields_dict.custom_join_teams_meeting.$wrapper
                .find("button")
                .off("click") 
                .on("click", () => {
                    window.open(frm.doc.custom_teams_meeting_url, "_blank");
                });
        }
        
        // 2. Handle Authentication Redirect
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get("teams_authentication_status") === "success") {
            frappe.show_alert({
                message: __('Teams token was successfully saved after login.'),
                indicator: 'green'
            });
            
            // Clean the URL like a pro
            const cleanURL = new URL(window.location.href);
            cleanURL.searchParams.delete('teams_authentication_status');
            window.history.replaceState({}, document.title, cleanURL.pathname);
        }

        // 3. Document Action Buttons (Only for saved docs)
        if (!frm.doc.__islocal) {
            
            // Create Teams Chat
            frm.add_custom_button(__('Create Teams Chat'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.create_group_chat_for_doc",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: (r) => {
                        if (r.message?.chat_id) {
                            frappe.msgprint(__("Teams chat created and linked to document."));
                            frm.reload_doc();
                        } else if (r.message?.login_url) {
                            window.location.href = r.message.login_url;
                        }
                    }
                });
            }, __("Teams"));

            // Open Teams Chat (Refactored to use Frappe Dialog!)
            frm.add_custom_button(__('Open Teams Chat'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.get_local_chat_messages",
                    args: { chat_id: frm.doc.custom_teams_chat_id },
                    callback: (r) => {
                        let messages = r.message || [];
                        
                        // Let's use the framework, Chandler!
                        let chat_html = messages.map(m => `
                            <div style="padding: 10px; border-bottom: 1px solid var(--border-color);">
                                <strong style="color: var(--text-color);">${m.sender_display || m.sender_id}</strong> 
                                <span style="color: var(--text-muted); font-size: 0.85em; margin-left: 8px;">${m.created_at || ''}</span>
                                <div style="margin-top: 6px; color: var(--text-light);">${m.body}</div>
                            </div>
                        `).join('');

                        if (!chat_html) chat_html = `<div class="text-muted p-4 text-center">No messages found.</div>`;

                        let d = new frappe.ui.Dialog({
                            title: __('Teams Chat History'),
                            fields: [
                                {
                                    fieldname: 'chat_container',
                                    fieldtype: 'HTML',
                                    options: `<div style="max-height: 400px; overflow-y: auto;">${chat_html}</div>`
                                }
                            ],
                            size: 'large'
                        });
                        
                        d.show();
                    }
                });
            }, __("Teams"));

            // Send Teams Message
            frm.add_custom_button(__('Send Teams Message'), () => {
                frappe.prompt([
                    {fieldname: 'message', fieldtype: 'Small Text', label: __('Message'), reqd: 1}
                ], (vals) => {
                    frappe.call({
                        method: "erpnext_teams_integration.api.chat.send_message_to_chat",
                        args: { 
                            chat_id: frm.doc.custom_teams_chat_id, 
                            message: vals.message, 
                            docname: frm.doc.name, 
                            doctype: frm.doc.doctype 
                        },
                        callback: () => {
                            frappe.show_alert({message: __('Message sent to Teams'), indicator: 'green'});
                            frm.reload_doc();
                        }
                    });
                }, __("Send Teams Message"), __("Send"));
            }, __("Teams"));

            // Post to Channel (UX Warning: Users hate typing IDs!)
            frm.add_custom_button(__('Post to Channel'), () => {
                frappe.prompt([
                    {fieldname: 'team_id', fieldtype: 'Data', label: __('Team ID'), reqd: 1},
                    {fieldname: 'channel_id', fieldtype: 'Data', label: __('Channel ID'), reqd: 1},
                    {fieldname: 'message', fieldtype: 'Small Text', label: __('Message'), reqd: 1}
                ], (vals) => {
                    frappe.call({
                        method: "erpnext_teams_integration.api.chat.post_message_to_channel",
                        args: { 
                            team_id: vals.team_id, 
                            channel_id: vals.channel_id, 
                            message: vals.message, 
                            docname: frm.doc.name, 
                            doctype: frm.doc.doctype 
                        },
                        callback: () => frappe.show_alert({message: __('Posted to channel'), indicator: 'blue'})
                    });
                }, __("Post to Channel"), __("Post"));
            }, __("Teams"));

            // Create Teams Meeting
            frm.add_custom_button(__('Create Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.create_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: (r) => {
                        if (r.message) {
                            let msg = typeof r.message === "string" ? r.message : r.message.message;
                            if (msg) frappe.msgprint(msg);
                        } 
                        if (r.message?.login_url) {
                            window.location.href = r.message.login_url;
                            return; // Stop execution if redirecting
                        }
                        frm.reload_doc();
                    }
                });
            }, __("Teams"));

            // Reschedule Teams Meeting
            frm.add_custom_button(__('Reschedule Teams Meeting'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.meetings.reschedule_meeting",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        if (r.message) {
                            // If it's an object, show the .message field
                            let msg = (typeof r.message === "string") ? r.message : r.message.message;
                            frappe.msgprint(msg);
                        } else if (r.message && r.message.login_url) {
                            // Redirect to MS login if required
                            window.location.href = r.message.login_url;
                        }
                        frm.reload_doc();
                    }
                });
            }, __("Teams"));

            

            // Sync Now
            frm.add_custom_button(__('Sync Now'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.sync_all_conversations",
                    args: frm.doc.custom_teams_chat_id ? { chat_id: frm.doc.custom_teams_chat_id } : {},
                    callback: (r) => {
                        if (!r.exc) frappe.show_alert({message: __('Chats synced successfully.'), indicator: 'green'});
                    }
                });
            }, __("Teams"));
        }
    }
});