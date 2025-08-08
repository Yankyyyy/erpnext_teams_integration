(function() {
    frappe.ui.form.on(cur_frm.doctype, {
        refresh: function(frm) {
            try {
                frappe.call({
                    method: "erpnext_teams_integration.api.settings.get_enabled_doctypes",
                    callback: function(r) {
                        var enabled = r.message || [];
                        if (enabled.indexOf(frm.doctype) !== -1) {
                            frm.add_custom_button("Open Teams Chat", function() {
                                frappe.call({
                                    method: "erpnext_teams_integration.api.chat.get_local_chat_messages",
                                    args: { chat_id: frm.doc.teams_chat_id },
                                    callback: function(res) {
                                        var messages = res.message || [];
                                        window.teamsChatModal = window.teamsChatModal || createTeamsModal();
                                        populateTeamsModal(messages);
                                        window.teamsChatModal.show();
                                    }
                                });
                            }, "Teams");
                            frm.add_custom_button("Send Teams Message", function() {
                                frappe.prompt([{fieldname:'message', fieldtype:'Small Text', label:'Message', reqd:1}], function(vals) {
                                    frappe.call({
                                        method: "erpnext_teams_integration.api.chat.send_message_to_chat",
                                        args: { chat_id: frm.doc.teams_chat_id, message: vals.message, docname: frm.doc.name, doctype: frm.doctype},
                                        callback: function(r) { frappe.msgprint("Message sent"); frm.reload_doc(); }
                                    });
                                }, "Send Teams Message", "Send");
                            }, "Teams");
                            frm.add_custom_button("Post to Channel", function() {
                                frappe.prompt([
                                    {fieldname:'team_id', fieldtype:'Data', label:'Team ID', reqd:1},
                                    {fieldname:'channel_id', fieldtype:'Data', label:'Channel ID', reqd:1},
                                    {fieldname:'message', fieldtype:'Small Text', label:'Message', reqd:1}
                                ], function(vals) {
                                    frappe.call({
                                        method: "erpnext_teams_integration.api.chat.post_message_to_channel",
                                        args: { team_id: vals.team_id, channel_id: vals.channel_id, message: vals.message, docname: frm.doc.name, doctype: frm.doctype},
                                        callback: function(r) { frappe.msgprint("Posted to channel"); }
                                    });
                                }, "Post to Channel", "Post");
                            }, "Teams");
                        }
                    }
                });
            } catch (e) { console.error(e); }
        }
    });
    function createTeamsModal() {
        var wrapper = document.createElement('div');
        wrapper.id = 'teams-chat-modal';
        wrapper.style = 'position:fixed;left:0;top:0;width:100%;height:100%;background:rgba(0,0,0,0.4);display:none;align-items:center;justify-content:center;z-index:9999;';
        var inner = document.createElement('div');
        inner.style = 'background:white;width:80%;max-width:900px;border-radius:8px;padding:16px;max-height:80%;overflow:auto;';
        inner.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px"><h4>Teams Chat</h4><button id="teams-close">Close</button></div><div id="teams-chat-contents"></div>';
        wrapper.appendChild(inner); document.body.appendChild(wrapper);
        wrapper.querySelector('#teams-close').addEventListener('click', function(){ wrapper.style.display='none'; });
        return { show: function(){ wrapper.style.display='flex'; }, hide: function(){ wrapper.style.display='none'; } };
    }
    function populateTeamsModal(messages) {
        var container = document.getElementById('teams-chat-contents');
        if(!container) return;
        container.innerHTML = '';
        messages.forEach(function(m){
            var el = document.createElement('div');
            el.style = 'padding:8px;border-bottom:1px solid #eee';
            el.innerHTML = '<b>'+(m.sender_display||m.sender_id)+'</b> <small style="color:#666">'+(m.created_at||'')+'</small><div style="margin-top:6px">'+m.body+'</div>';
            container.appendChild(el);
        });
    }
})();