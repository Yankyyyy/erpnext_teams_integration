frappe.ui.form.on('Teams Settings', {
    refresh: function(frm) {
        if (!frm.is_new()) {
            
            // Main authentication button
            frm.add_custom_button(__('Authenticate with Teams'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.helpers.get_login_url",
                    args: { docname: frm.doc.name },
                    callback: function(r) {
                        if (r.message) {
                            window.location.href = r.message; // redirect to MS login

                            frappe.msgprint({
                                title: __('Authentication Started'),
                                message: __('Please complete the authentication in the new window. This page will automatically refresh when authentication is complete.'),
                                indicator: 'blue'
                            });
                        }
                    }
                });
            }).addClass("btn-primary");

            // Test connection button
            frm.add_custom_button(__('Test Connection'), function() {
                frappe.show_alert({
                    message: __('Testing connection...'),
                    indicator: 'blue'
                });
                
                frappe.call({
                    method: "erpnext_teams_integration.api.settings.test_teams_connection",
                    callback: function(r) {
                        if (r.message) {
                            if (r.message.success) {
                                const user_info = r.message.user_info;
                                const permissions = r.message.permissions;
                                
                                let message = `Connected as: <b>${user_info.name}</b> (${user_info.email})<br>`;
                                message += `<br><b>Permissions:</b><br>`;
                                message += `• Chat Access: ${permissions.chats ? '✅' : '❌'}<br>`;
                                message += `• Meetings Access: ${permissions.meetings ? '✅' : '❌'}`;
                                
                                frappe.msgprint({
                                    title: __('Connection Test Successful'),
                                    message: message,
                                    indicator: 'green'
                                });
                            } else {
                                frappe.msgprint({
                                    title: __('Connection Test Failed'),
                                    message: r.message.message || 'Unknown error',
                                    indicator: 'red'
                                });
                            }
                        }
                    }
                });
            }, __('More Actions'));

            // Sync conversations button
            frm.add_custom_button(__('Sync All Conversations'), function() {
                frappe.confirm(
                    __('This will fetch recent messages from all your Teams chats. This may take a while. Continue?'),
                    function() {
                        frappe.show_alert({
                            message: __('Syncing conversations...'),
                            indicator: 'blue'
                        });
                        
                        frappe.call({
                            method: "erpnext_teams_integration.api.chat.sync_all_conversations",
                            callback: function(r) {
                                if (!r.exc && r.message) {
                                    frappe.msgprint({
                                        title: __('Sync Complete'),
                                        message: `Synced ${r.message.synced} conversations successfully. ${r.message.errors} errors occurred.`,
                                        indicator: 'green'
                                    });
                                } else {
                                    frappe.msgprint({
                                        title: __('Sync Failed'),
                                        message: __('Failed to sync conversations. Check error logs.'),
                                        indicator: 'red'
                                    });
                                }
                            }
                        });
                    }
                );
            }, __('Sync Actions'));

            // Bulk sync Azure IDs button
            frm.add_custom_button(__('Sync Azure IDs'), function() {
                frappe.confirm(
                    __('This will fetch all users from your Microsoft tenant and update their Azure Object IDs. Continue?'),
                    function() {
                        frappe.show_alert({
                            message: __('Syncing Azure IDs...'),
                            indicator: 'blue'
                        });
                        
                        frappe.call({
                            method: "erpnext_teams_integration.api.settings.bulk_sync_azure_ids",
                            callback: function(r) {
                                if (!r.exc) {
                                    // Update owner Azure ID if it was synced
                                    if (frm.doc.azure_owner_email_id) {
                                        frappe.db.get_value('User', frm.doc.azure_owner_email_id, 'azure_object_id')
                                            .then(res => {
                                                if (res && res.message && res.message.azure_object_id) {
                                                    frm.set_value('owner_azure_object_id', res.message.azure_object_id);
                                                    frm.save();
                                                }
                                            });
                                    }
                                    
                                    frappe.msgprint({
                                        title: __('Sync Complete'),
                                        message: r.message || __('Azure IDs synced successfully.'),
                                        indicator: 'green'
                                    });
                                }
                            }
                        });
                    }
                );
            }, __('Sync Actions'));

            // Get statistics button
            frm.add_custom_button(__('View Statistics'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.settings.get_teams_statistics",
                    callback: function(r) {
                        if (r.message) {
                            const stats = r.message;
                            let message = `<div style="font-family: monospace;">
                                <b>Teams Integration Statistics</b><br><br>
                                📊 <b>Messages:</b><br>
                                • Total Messages: ${stats.total_messages}<br>
                                • Inbound: ${stats.inbound_messages}<br>
                                • Outbound: ${stats.outbound_messages}<br><br>
                                
                                💬 <b>Conversations:</b><br>
                                • Total Conversations: ${stats.total_conversations}<br>
                                • Unique Chats: ${stats.unique_chats}<br><br>
                                
                                👥 <b>Users:</b><br>
                                • Users with Azure ID: ${stats.users_with_azure_id}<br>
                            </div>`;
                            
                            if (stats.recent_activity && stats.recent_activity.length > 0) {
                                message += `<br><b>Recent Activity (Last 7 days):</b><br>`;
                                stats.recent_activity.forEach(activity => {
                                    message += `• ${activity.date}: ${activity.count} messages<br>`;
                                });
                            }
                            
                            frappe.msgprint({
                                title: __('Teams Integration Statistics'),
                                message: message,
                                indicator: 'blue',
                                wide: true
                            });
                        }
                    }
                });
            }, __('More Actions'));

            // Validate configuration button
            frm.add_custom_button(__('Validate Configuration'), function() {
                frappe.call({
                    method: "erpnext_teams_integration.api.settings.validate_configuration",
                    callback: function(r) {
                        if (r.message) {
                            const result = r.message;
                            
                            if (result.valid) {
                                frappe.msgprint({
                                    title: __('Configuration Valid'),
                                    message: __('Your Teams integration configuration is correct and ready to use!'),
                                    indicator: 'green'
                                });
                            } else {
                                let message = '<b>Configuration Issues Found:</b><br>';
                                result.issues.forEach(issue => {
                                    message += `• ${issue}<br>`;
                                });
                                
                                frappe.msgprint({
                                    title: __('Configuration Issues'),
                                    message: message,
                                    indicator: 'red'
                                });
                            }
                        }
                    }
                });
            }, __('More Actions'));

            // // Add advanced options dropdown
            // frm.add_custom_button(__('Advanced Options'), function() {
            //     // This will show a dropdown with advanced options
            // }, __('More Actions'));

            // Reset integration (dangerous operation)
            frm.add_custom_button(__('Reset Integration'), function() {
                frappe.warn(
                    __('Reset Teams Integration'),
                    __('This will clear all authentication tokens and conversation data. This action cannot be undone. Are you absolutely sure?'),
                    function() {
                        frappe.call({
                            method: "erpnext_teams_integration.api.settings.reset_integration",
                            callback: function(r) {
                                if (!r.exc) {
                                    frm.reload_doc();
                                    frappe.msgprint({
                                        title: __('Integration Reset'),
                                        message: r.message,
                                        indicator: 'orange'
                                    });
                                }
                            }
                        });
                    }
                );
            }, __('More Actions'));

            // Cleanup old messages
            frm.add_custom_button(__('Cleanup Old Messages'), function() {
                frappe.prompt([{
                    'fieldname': 'days',
                    'label': __('Delete messages older than (days)'),
                    'fieldtype': 'Int',
                    'default': 30,
                    'reqd': 1,
                    'description': __('Messages older than this many days will be permanently deleted')
                }], function(values) {
                    frappe.confirm(
                        __(`This will permanently delete all Teams messages older than ${values.days} days. Continue?`),
                        function() {
                            frappe.call({
                                method: "erpnext_teams_integration.api.settings.cleanup_old_messages",
                                args: { days: values.days },
                                callback: function(r) {
                                    if (!r.exc) {
                                        frappe.msgprint({
                                            title: __('Cleanup Complete'),
                                            message: r.message,
                                            indicator: 'green'
                                        });
                                    }
                                }
                            });
                        }
                    );
                }, __('Cleanup Messages'), __('Delete'));
            }, __('More Actions'));

            // Show authentication status indicator
            if (frm.doc.access_token) {
                frm.dashboard.add_indicator(__('Authenticated'), 'green');
                
                // Check token expiry
                if (frm.doc.token_expiry) {
                    const expiry = new Date(frm.doc.token_expiry);
                    const now = new Date();
                    const hoursUntilExpiry = (expiry - now) / (1000 * 60 * 60);
                    
                    if (hoursUntilExpiry <= 0) {
                        // Token already expired
                        frm.dashboard.add_indicator(__('Token Expired'), 'red');
                    } else if (hoursUntilExpiry < 1) {
                        // Token expiring within next hour
                        frm.dashboard.add_indicator(__('Token Expires Soon'), 'yellow');
                    }
                }
            } else {
                frm.dashboard.add_indicator(__('Not Authenticated'), 'red');
            }

            // Add help section
            frm.dashboard.add_section(`
                <div style="padding: 10px; background: #f8f9fa; border-radius: 5px; margin: 10px 0;">
                    <h5>Setup Instructions:</h5>
                    <ol>
                        <li>Configure your Microsoft Azure App registration details above</li>
                        <li>Click "Authenticate with Teams" to authorize the app</li>
                        <li>Use "Test Connection" to verify everything is working</li>
                        <li>Enable doctypes in the "Enabled Doctypes" table</li>
                        <li>Use "Sync Azure IDs" to link Frappe users with Microsoft accounts</li>
                    </ol>
                    <p><strong>Need help?</strong> Check the <a href="https://docs.microsoft.com/en-us/graph/auth-register-app-v2" target="_blank">Microsoft Graph documentation</a> for app registration guidance.</p>
                </div>
            `);
        }

        // Handle URL parameters (e.g., from OAuth redirect)
        const urlParams = new URLSearchParams(window.location.search);
        const authStatus = urlParams.get('teams_authentication_status');
        
        if (authStatus === 'success') {
            frappe.show_alert({
                message: __('Teams authentication successful!'),
                indicator: 'green'
            }, 5);
            
            // Clean up URL
            const cleanURL = new URL(window.location.href);
            cleanURL.searchParams.delete('teams_authentication_status');
            window.history.replaceState({}, document.title, cleanURL.pathname);
            
            // Reload form to show updated status
            setTimeout(() => {
                frm.reload_doc();
            }, 1000);
            
        } else if (authStatus === 'error') {
            frappe.show_alert({
                message: __('Teams authentication failed. Please try again.'),
                indicator: 'red'
            }, 10);
            
            // Clean up URL
            const cleanURL = new URL(window.location.href);
            cleanURL.searchParams.delete('teams_authentication_status');
            window.history.replaceState({}, document.title, cleanURL.pathname);
        }
    },

    // Validate redirect URI format
    redirect_uri: function(frm) {
        if (frm.doc.redirect_uri) {
            if (!frm.doc.redirect_uri.startsWith('http://') && !frm.doc.redirect_uri.startsWith('https://')) {
                frappe.msgprint({
                    title: __('Invalid Redirect URI'),
                    message: __('Redirect URI must start with http:// or https://'),
                    indicator: 'red'
                });
            }
        }
    },

    // Validate tenant ID format
    tenant_id: function(frm) {
        if (frm.doc.tenant_id) {
            const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
            if (!guidPattern.test(frm.doc.tenant_id)) {
                frappe.msgprint({
                    title: __('Invalid Tenant ID'),
                    message: __('Tenant ID should be in GUID format (e.g., 12345678-1234-1234-1234-123456789012)'),
                    indicator: 'orange'
                });
            }
        }
    },

    // Auto-populate redirect URI
    onload: function(frm) {
        if (frm.is_new() || !frm.doc.redirect_uri) {
            const currentSite = window.location.origin;
            const defaultRedirectUri = `${currentSite}/api/method/erpnext_teams_integration.api.auth.callback`;
            frm.set_value('redirect_uri', defaultRedirectUri);
        }
    }
});