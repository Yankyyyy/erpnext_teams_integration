# import frappe
# def after_install():
#     if not frappe.db.exists("Custom Field", {"dt":"User", "fieldname":"azure_object_id"}):
#         frappe.get_doc({ "doctype":"Custom Field", "dt":"User", "fieldname":"azure_object_id", "label":"Azure Object ID", "fieldtype":"Data", "insert_after":"email" }).insert(ignore_permissions=True)
#     frappe.db.commit()

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_field


def after_install():
    """Post-installation setup for ERPNext Teams Integration"""
    try:
        # Create custom field for Azure Object ID in User doctype
        create_azure_object_id_field()
        
        # Create Teams Settings singleton if it doesn't exist
        create_teams_settings()
        
        # Set up default permissions
        setup_permissions()
        
        # Create indexes for better performance
        create_database_indexes()
        
        # Show installation success message
        print("ERPNext Teams Integration installed successfully!")
        print("Next steps:")
        print("   1. Go to Teams Settings")
        print("   2. Configure your Microsoft Azure app credentials")
        print("   3. Authenticate with Microsoft Teams")
        print("   4. Enable the doctypes you want to integrate")
        
        frappe.db.commit()
        
    except Exception as e:
        frappe.log_error(f"Installation error: {str(e)}", "Teams Integration Install Error")
        print(f"Installation error: {str(e)}")
        raise


def create_azure_object_id_field():
    """Create Azure Object ID custom field in User doctype"""
    try:
        # Check if custom field already exists
        if not frappe.db.exists("Custom Field", {"dt": "User", "fieldname": "azure_object_id"}):
            custom_field = {
                "doctype": "Custom Field",
                "dt": "User",
                "fieldname": "azure_object_id",
                "label": "Azure Object ID",
                "fieldtype": "Data",
                "insert_after": "email",
                "read_only": 1,
                "no_copy": 1,
                "hidden": 0,
                "description": "Microsoft Azure Active Directory Object ID for Teams integration"
            }
            
            create_custom_field("User", custom_field)
            print("Created Azure Object ID field in User doctype")
        else:
            print("ℹAzure Object ID field already exists in User doctype")
            
    except Exception as e:
        frappe.log_error(f"Error creating Azure Object ID field: {str(e)}", "Teams Install Field Error")
        print(f"Warning: Could not create Azure Object ID field: {str(e)}")


def create_teams_settings():
    """Create Teams Settings singleton document"""
    try:
        if not frappe.db.exists("Teams Settings", "Teams Settings"):
            # Get current site URL for default redirect URI
            site_url = frappe.utils.get_url()
            default_redirect_uri = f"{site_url}/api/method/erpnext_teams_integration.api.auth.callback"
            
            settings_doc = frappe.get_doc({
                "doctype": "Teams Settings",
                "redirect_uri": default_redirect_uri,
                "enabled_doctypes": [
                    {"doctype_name": "Event"},
                    {"doctype_name": "Project"}
                ]
            })
            settings_doc.insert(ignore_permissions=True)
            print("Created Teams Settings document with default configuration")
        else:
            print("Teams Settings document already exists")
            
    except Exception as e:
        frappe.log_error(f"Error creating Teams Settings: {str(e)}", "Teams Install Settings Error")
        print(f"Warning: Could not create Teams Settings: {str(e)}")


def setup_permissions():
    """Set up default permissions for Teams doctypes"""
    try:
        # Teams Settings permissions
        ensure_doctype_permissions("Teams Settings", [
            {"role": "System Manager", "read": 1, "write": 1, "create": 1, "delete": 1},
            {"role": "Administrator", "read": 1, "write": 1, "create": 1, "delete": 1}
        ])
        
        # Teams Conversation permissions
        ensure_doctype_permissions("Teams Conversation", [
            {"role": "System Manager", "read": 1, "write": 1, "create": 1, "delete": 1},
            {"role": "All", "read": 1, "write": 0, "create": 0, "delete": 0}
        ])
        
        # Teams Chat Message permissions
        ensure_doctype_permissions("Teams Chat Message", [
            {"role": "System Manager", "read": 1, "write": 1, "create": 1, "delete": 1},
            {"role": "All", "read": 1, "write": 0, "create": 1, "delete": 0}
        ])
        
        print("Set up default permissions for Teams doctypes")
        
    except Exception as e:
        frappe.log_error(f"Error setting up permissions: {str(e)}", "Teams Install Permissions Error")
        print(f"Warning: Could not set up all permissions: {str(e)}")


def ensure_doctype_permissions(doctype, permissions):
    """Ensure specific permissions exist for a doctype"""
    try:
        for perm in permissions:
            # Check if permission already exists
            existing_perm = frappe.db.exists("DocPerm", {
                "parent": doctype,
                "role": perm["role"]
            })
            
            if not existing_perm:
                # Create new permission
                perm_doc = frappe.get_doc({
                    "doctype": "DocPerm",
                    "parent": doctype,
                    "parentfield": "permissions",
                    "parenttype": "DocType",
                    **perm
                })
                perm_doc.insert(ignore_permissions=True)
                
    except Exception as e:
        frappe.log_error(f"Error ensuring permissions for {doctype}: {str(e)}", "Teams Permission Error")


def create_database_indexes():
    """Create database indexes for better performance"""
    try:
        # Index on Teams Chat Message for faster queries
        indexes = [
            {
                "table": "tabTeams Chat Message",
                "columns": ["chat_id", "created_at"],
                "name": "idx_teams_chat_message_chat_created"
            },
            {
                "table": "tabTeams Chat Message",
                "columns": ["message_id"],
                "name": "idx_teams_chat_message_id",
                "unique": True
            },
            {
                "table": "tabTeams Chat Message",
                "columns": ["direction", "created_at"],
                "name": "idx_teams_chat_message_direction_created"
            },
            {
                "table": "tabTeams Conversation",
                "columns": ["chat_id"],
                "name": "idx_teams_conversation_chat_id",
                "unique": True
            },
            {
                "table": "tabUser",
                "columns": ["azure_object_id"],
                "name": "idx_user_azure_object_id"
            }
        ]
        
        for index in indexes:
            try:
                columns_str = ", ".join([f"`{col}`" for col in index["columns"]])
                unique_str = "UNIQUE" if index.get("unique") else ""
                
                # Check if index already exists
                check_sql = f"""
                    SELECT COUNT(*) as count FROM information_schema.statistics 
                    WHERE table_schema = DATABASE() 
                    AND table_name = '{index["table"]}' 
                    AND index_name = '{index["name"]}'
                """
                result = frappe.db.sql(check_sql, as_dict=True)
                
                if result[0]["count"] == 0:
                    create_sql = f"""
                        CREATE {unique_str} INDEX `{index["name"]}` 
                        ON `{index["table"]}` ({columns_str})
                    """
                    frappe.db.sql(create_sql)
                    print(f"Created index {index['name']}")
                    
            except Exception as idx_error:
                # Log but don't fail installation for index errors
                frappe.log_error(f"Error creating index {index['name']}: {str(idx_error)}", "Teams Index Error")
                print(f"Could not create index {index['name']}: {str(idx_error)}")
        
    except Exception as e:
        frappe.log_error(f"Error creating database indexes: {str(e)}", "Teams Install Index Error")
        print(f"Warning: Could not create all database indexes: {str(e)}")


def before_uninstall():
    """Cleanup before app uninstall"""
    try:
        print("Cleaning up Teams Integration...")
        
        # Optionally backup data before cleanup
        backup_teams_data()
        
        # Remove custom fields (optional - commented out to preserve data)
        # remove_custom_fields()
        
        # Remove indexes
        remove_database_indexes()
        
        print("Teams Integration cleanup completed")
        
    except Exception as e:
        frappe.log_error(f"Uninstall cleanup error: {str(e)}", "Teams Uninstall Error")
        print(f"Warning: Cleanup error: {str(e)}")


def backup_teams_data():
    """Backup Teams data before uninstall"""
    try:
        import json
        from frappe.utils import now
        
        # Get all Teams data
        conversations = frappe.get_all("Teams Conversation", fields="*")
        messages = frappe.get_all("Teams Chat Message", fields="*")
        
        backup_data = {
            "conversations": conversations,
            "messages": messages,
            "backup_time": now(),
            "version": "1.0"
        }
        
        # Save to file
        backup_path = f"/tmp/teams_integration_backup_{now().replace(' ', '_').replace(':', '-')}.json"
        with open(backup_path, 'w') as f:
            json.dump(backup_data, f, indent=2, default=str)
        
        print(f"Teams data backed up to {backup_path}")
        
    except Exception as e:
        print(f"Could not backup Teams data: {str(e)}")


def remove_custom_fields():
    """Remove custom fields created by the app"""
    try:
        # Remove Azure Object ID field
        if frappe.db.exists("Custom Field", {"dt": "User", "fieldname": "azure_object_id"}):
            frappe.delete_doc("Custom Field", {"dt": "User", "fieldname": "azure_object_id"})
            print("Removed Azure Object ID custom field")
            
    except Exception as e:
        print(f"Could not remove custom fields: {str(e)}")


def remove_database_indexes():
    """Remove database indexes created by the app"""
    try:
        indexes_to_remove = [
            "idx_teams_chat_message_chat_created",
            "idx_teams_chat_message_id", 
            "idx_teams_chat_message_direction_created",
            "idx_teams_conversation_chat_id",
            "idx_user_azure_object_id"
        ]
        
        for index_name in indexes_to_remove:
            try:
                # Find which table the index belongs to
                check_sql = f"""
                    SELECT table_name FROM information_schema.statistics 
                    WHERE table_schema = DATABASE() AND index_name = '{index_name}'
                    LIMIT 1
                """
                result = frappe.db.sql(check_sql, as_dict=True)
                
                if result:
                    table_name = result[0]["table_name"]
                    drop_sql = f"DROP INDEX `{index_name}` ON `{table_name}`"
                    frappe.db.sql(drop_sql)
                    print(f"Removed index {index_name}")
                    
            except Exception as idx_error:
                print(f"Could not remove index {index_name}: {str(idx_error)}")
                
    except Exception as e:
        print(f"Error removing indexes: {str(e)}")