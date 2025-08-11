frappe.ui.form.on("Project", {
    refresh(frm) {
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get("teams_authentication_status") === "success") {
            frappe.msgprint({
                title: "Token Fetched Successfully",
                message: "Teams token was successfully saved after login.",
                indicator: 'green'
            });
            const cleanURL = new URL(window.location.href);
            cleanURL.searchParams.delete('teams_authentication_status');
            window.history.replaceState({}, document.title, cleanURL.pathname);
        }

        if (!frm.doc.__islocal) {
            // Create a dropdown called "Teams"
            frm.add_custom_button(__('Create Teams Chat'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.create_group_chat_for_doc",
                    args: { docname: frm.doc.name, doctype: frm.doc.doctype },
                    callback: function(r) {
                        if (r.message && r.message.chat_id) {
                            frappe.msgprint("Teams chat created and linked to document.");
                            frm.reload_doc();
                        } else if (r.message && r.message.login_url) {
                            window.location.href = r.message.login_url;
                        }
                    }
                });
            }, __("Teams"));

            frm.add_custom_button(__('Open Teams Chat'), () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.get_local_chat_messages",
                    args: { chat_id: frm.doc.custom_teams_chat_id },
                    callback: function(r) {
                        var messages = r.message || [];
                        window.teamsChatModal = window.teamsChatModal || (function(){
                            var w=document.createElement('div');
                            w.id='teams-chat-modal';
                            w.style='position:fixed;left:0;top:0;width:100%;height:100%;background:rgba(0,0,0,0.4);display:flex;align-items:center;justify-content:center;z-index:9999;';
                            var inner=document.createElement('div');
                            inner.style='background:white;width:80%;max-width:900px;border-radius:8px;padding:16px;max-height:80%;overflow:auto;';
                            inner.innerHTML='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px"><h4>Teams Chat</h4><button id="teams-close">Close</button></div><div id="teams-chat-contents"></div>';
                            w.appendChild(inner);
                            document.body.appendChild(w);
                            w.querySelector('#teams-close').addEventListener('click',function(){w.style.display='none'});
                            return {show:function(){w.style.display='flex'},hide:function(){w.style.display='none'}};
                        })();
                        var container = document.getElementById('teams-chat-contents');
                        container.innerHTML='';
                        messages.forEach(function(m){
                            var el=document.createElement('div');
                            el.style='padding:8px;border-bottom:1px solid #eee';
                            el.innerHTML = '<b>'+(m.sender_display||m.sender_id)+'</b> <small style="color:#666">'+(m.created_at||'')+'</small><div style="margin-top:6px">'+m.body+'</div>';
                            container.appendChild(el);
                        });
                        window.teamsChatModal.show();
                    }
                });
            }, __("Teams"));

            frm.add_custom_button(__('Send Teams Message'), () => {
                frappe.prompt([{fieldname:'message', fieldtype:'Small Text', label:'Message', reqd:1}], function(vals) {
                    frappe.call({
                        method: "erpnext_teams_integration.api.chat.send_message_to_chat",
                        args: { chat_id: frm.doc.custom_teams_chat_id, message: vals.message, docname: frm.doc.name, doctype: frm.doc.doctype },
                        callback: function() {
                            frappe.msgprint('Message sent');
                            frm.reload_doc();
                        }
                    });
                }, "Send Teams Message", "Send");
            }, __("Teams"));

            frm.add_custom_button(__('Post to Channel'), () => {
                frappe.prompt([
                    {fieldname:'team_id', fieldtype:'Data', label:'Team ID', reqd:1},
                    {fieldname:'channel_id', fieldtype:'Data', label:'Channel ID', reqd:1},
                    {fieldname:'message', fieldtype:'Small Text', label:'Message', reqd:1}
                ], function(vals) {
                    frappe.call({
                        method: "erpnext_teams_integration.api.chat.post_message_to_channel",
                        args: { team_id: vals.team_id, channel_id: vals.channel_id, message: vals.message, docname: frm.doc.name, doctype: frm.doc.doctype },
                        callback: function() { frappe.msgprint('Posted to channel'); }
                    });
                }, "Post to Channel", "Post");
            }, __("Teams"));

            // New "Sync Now" button inside Teams dropdown
            frm.add_custom_button(__('Sync Now'), () => {
                let args = {};
                if (frm.doc.custom_teams_chat_id) {
                    args.chat_id = frm.doc.custom_teams_chat_id;
                }

                frappe.call({
                    method: "erpnext_teams_integration.api.chat.sync_all_conversations",
                    args: args,
                    callback: function(r) {
                        if (!r.exc) {
                            frappe.msgprint(__('Chats synced successfully.'));
                        }
                    }
                });
            }, __("Teams"));
        }
    }
});