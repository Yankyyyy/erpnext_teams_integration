frappe.ui.form.on("Event", {
    refresh(frm) {
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get("teams_authentication_status") === "success") {
            frappe.msgprint({title: "Token Fetched Successfully", message: "Teams token was successfully saved after login.", indicator: 'green'});
            const cleanURL = new URL(window.location.href);
            cleanURL.searchParams.delete('teams_authentication_status');
            window.history.replaceState({}, document.title, cleanURL.pathname);
        }
        if (!frm.doc.__islocal) {
            frm.add_custom_button("Create Teams Chat", () => {
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
            });
            frm.add_custom_button("Open Teams Chat", () => {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.get_local_chat_messages",
                    args: { chat_id: frm.doc.custom_teams_chat_id },
                    callback: function(r) {
                        var messages = r.message || [];
                        window.teamsChatModal = window.teamsChatModal || (function(){ var w=document.createElement('div'); w.id='teams-chat-modal'; w.style='position:fixed;left:0;top:0;width:100%;height:100%;background:rgba(0,0,0,0.4);display:flex;align-items:center;justify-content:center;z-index:9999;'; var inner=document.createElement('div'); inner.style='background:white;width:80%;max-width:900px;border-radius:8px;padding:16px;max-height:80%;overflow:auto;'; inner.innerHTML='<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px"><h4>Teams Chat</h4><button id="teams-close">Close</button></div><div id="teams-chat-contents"></div>'; w.appendChild(inner); document.body.appendChild(w); w.querySelector('#teams-close').addEventListener('click',function(){w.style.display='none'}); return {show:function(){w.style.display='flex'},hide:function(){w.style.display='none'}}; })();
                        var container = document.getElementById('teams-chat-contents');
                        container.innerHTML='';
                        messages.forEach(function(m){ var el=document.createElement('div'); el.style='padding:8px;border-bottom:1px solid #eee'; el.innerHTML = '<b>'+(m.sender_display||m.sender_id)+'</b> <small style="color:#666">'+(m.created_at||'')+'</small><div style="margin-top:6px">'+m.body+'</div>'; container.appendChild(el); });
                        window.teamsChatModal.show();
                    }
                });
            });
            frm.add_custom_button("Send Teams Message", () => {
                frappe.prompt([{fieldname:'message', fieldtype:'Small Text', label:'Message', reqd:1}], function(vals) {
                    frappe.call({
                        method: "erpnext_teams_integration.api.chat.send_message_to_chat",
                        args: { chat_id: frm.doc.custom_teams_chat_id, message: vals.message, docname: frm.doc.name, doctype: frm.doc.doctype },
                        callback: function() { frappe.msgprint('Message sent'); frm.reload_doc(); }
                    });
                }, "Send Teams Message", "Send");
            });
            frm.add_custom_button("Post to Channel", () => {
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
            });
        }
    }
});

function createTeamsModal() {
    let wrapper = document.createElement('div');
    wrapper.id = 'teams-chat-modal';
    wrapper.style = 'position:fixed;left:0;top:0;width:100%;height:100%;background:rgba(0,0,0,0.4);display:none;align-items:center;justify-content:center;z-index:9999;';
    let inner = document.createElement('div');
    inner.style = 'background:white;width:80%;max-width:900px;border-radius:8px;padding:16px;max-height:80%;overflow:auto;';
    inner.innerHTML = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px"><h4>Teams Chat</h4><button id="teams-close">Close</button></div><div id="teams-chat-contents"></div>';
    wrapper.appendChild(inner);
    document.body.appendChild(wrapper);
    wrapper.querySelector('#teams-close').addEventListener('click', function(){ wrapper.style.display='none'; });
    return { show: function(){ wrapper.style.display='flex'; }, hide: function(){ wrapper.style.display='none'; } };
}

function populateTeamsModal(messages) {
    let container = document.getElementById('teams-chat-contents');
    if (!container) return;
    container.innerHTML = '';
    messages.forEach(function(m) {
        let el = document.createElement('div');
        el.style = 'padding:8px;border-bottom:1px solid #eee';
        el.innerHTML = `<b>${m.sender_display || m.sender_id}</b> <small style="color:#666">${m.created_at || ''}</small><div style="margin-top:6px">${m.body}</div>`;
        container.appendChild(el);
    });
}