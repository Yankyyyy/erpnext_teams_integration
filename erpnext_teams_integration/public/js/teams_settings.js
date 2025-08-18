frappe.ui.form.on('Teams Settings', {
    refresh: function(frm) {
        if (!frm.is_new()) {

            // Authenticate Button
            frm.add_custom_button(__('Authenticate with Teams'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.helpers.get_login_url",
                     args: { docname: frm.doc.name },
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

            // Bulk update azure object IDs
            frm.add_custom_button(__('Sync Azure IDs'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.settings.bulk_sync_azure_ids",
                    callback: function(r) {
                        if (!r.exc) {
                            if (r.message && frm.doc.azure_owner_email_id) {
                                frappe.db.get_value('User', frm.doc.azure_owner_email_id, 'azure_object_id')
                                    .then(res => {
                                        if (res && res.message) {
                                            frm.set_value('owner_azure_object_id', res.message.azure_object_id);
                                            frm.save();
                                        }
                                    });
                            }
                            frappe.msgprint(__(r.message || 'Azure IDs synced successfully.'));
                        }
                    }
                });
            });
        }
    }
});
