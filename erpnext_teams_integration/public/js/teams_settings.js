frappe.ui.form.on('Teams Settings', {
    refresh: function(frm) {
        if (!frm.is_new()) {

            // Authenticate Button
            frm.add_custom_button(__('Authenticate with Teams'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.helpers.get_login_url",
                    callback: function(r) {
                        if (r.message) {
                            window.location.href = r.message; // redirect to MS login
                        }
                    }
                });
            }).addClass("btn-primary");

            // Sync Now Button
            frm.add_custom_button(__('Sync Now'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.chat.sync_all_conversations",
                    callback: function(r) {
                        if (!r.exc) {
                            frappe.msgprint(__('Chats synced successfully.'));
                        }
                    }
                });
            });
        }
    }
});
