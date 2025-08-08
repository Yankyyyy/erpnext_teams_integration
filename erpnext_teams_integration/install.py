import frappe
def after_install():
    if not frappe.db.exists("Custom Field", {"dt":"User", "fieldname":"azure_object_id"}):
        frappe.get_doc({ "doctype":"Custom Field", "dt":"User", "fieldname":"azure_object_id", "label":"Azure Object ID", "fieldtype":"Data", "insert_after":"email" }).insert(ignore_permissions=True)
    frappe.db.commit()
