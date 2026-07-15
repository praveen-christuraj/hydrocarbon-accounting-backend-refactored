from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import Permission, RolePermission, Role, User
from app.schemas import PermissionCreate, PermissionResponse
from app.dependencies.auth import get_current_user_from_token
from app.dependencies.permissions import require_user_permission
from app.services.audit_service import create_audit_log
from app.utils.helpers import clean_optional_text
from app.utils.pagination import paginate_query

router = APIRouter(prefix="/permissions", tags=["Permissions"])


@router.get("")
def get_permissions(
    skip: int = Query(0, ge=0),
    limit: int = Query(50, ge=1, le=200),
    search: str | None = Query(None),
    module_name: str | None = Query(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_from_token),
):
    require_user_permission(current_user, "View Permission", db)
    query = db.query(Permission).order_by(Permission.id)
    if search:
        query = query.filter(Permission.permission_name.ilike(f"%{search}%"))
    if module_name:
        query = query.filter(Permission.module_name.ilike(f"%{module_name}%"))
    result = paginate_query(query, skip, limit)
    return {
        "items": [PermissionResponse.model_validate(p) for p in result["items"]],
        "total": result["total"],
        "skip": result["skip"],
        "limit": result["limit"],
        "has_more": result["has_more"],
    }


@router.post("", response_model=PermissionResponse)
def create_permission(
    permission: PermissionCreate,
    current_user: User = Depends(get_current_user_from_token),
    db: Session = Depends(get_db),
):
    require_user_permission(
        current_user,
        "Manage Permission",
        db,
    )

    existing_permission = (
        db.query(Permission)
        .filter(
            Permission.permission_name.ilike(permission.permission_name),
            Permission.module_name.ilike(permission.module_name),
        )
        .first()
    )

    if existing_permission:
        raise HTTPException(
            status_code=400,
            detail="Permission already exists for this module",
        )

    new_permission = Permission(
        permission_name=permission.permission_name.strip(),
        module_name=permission.module_name.strip(),
        description=clean_optional_text(permission.description),
        status=permission.status,
    )

    db.add(new_permission)
    db.flush()

    after_data = {
        "permission_name": new_permission.permission_name,
        "module_name": new_permission.module_name,
        "description": new_permission.description,
        "status": new_permission.status,
    }

    create_audit_log(
        db=db,
        module_name="Permission Master",
        action="Create Permission",
        current_user=current_user,
        entity_type="Permission",
        entity_id=new_permission.id,
        entity_label=f"{new_permission.module_name} - {new_permission.permission_name}",
        remarks="Permission created",
        request_path="/permissions",
        details={
            "after": after_data,
        },
    )

    db.commit()
    db.refresh(new_permission)

    return new_permission


@router.put("/{permission_id}", response_model=PermissionResponse)
def update_permission(
    permission_id: int,
    permission: PermissionCreate,
    current_user: User = Depends(get_current_user_from_token),
    db: Session = Depends(get_db),
):
    require_user_permission(
        current_user,
        "Manage Permission",
        db,
    )

    existing_permission = (
        db.query(Permission)
        .filter(Permission.id == permission_id)
        .first()
    )

    if not existing_permission:
        raise HTTPException(
            status_code=404,
            detail="Permission not found",
        )

    duplicate_permission = (
        db.query(Permission)
        .filter(
            Permission.permission_name.ilike(permission.permission_name),
            Permission.module_name.ilike(permission.module_name),
            Permission.id != permission_id,
        )
        .first()
    )

    if duplicate_permission:
        raise HTTPException(
            status_code=400,
            detail="Permission already exists for this module",
        )

    before_data = {
        "permission_name": existing_permission.permission_name,
        "module_name": existing_permission.module_name,
        "description": existing_permission.description,
        "status": existing_permission.status,
    }

    existing_permission.permission_name = permission.permission_name.strip()
    existing_permission.module_name = permission.module_name.strip()
    existing_permission.description = clean_optional_text(permission.description)
    existing_permission.status = permission.status

    after_data = {
        "permission_name": existing_permission.permission_name,
        "module_name": existing_permission.module_name,
        "description": existing_permission.description,
        "status": existing_permission.status,
    }

    create_audit_log(
        db=db,
        module_name="Permission Master",
        action="Update Permission",
        current_user=current_user,
        entity_type="Permission",
        entity_id=existing_permission.id,
        entity_label=f"{existing_permission.module_name} - {existing_permission.permission_name}",
        remarks="Permission updated",
        request_path=f"/permissions/{permission_id}",
        details={
            "before": before_data,
            "after": after_data,
        },
    )

    db.commit()
    db.refresh(existing_permission)

    return existing_permission


@router.delete("/{permission_id}")
def delete_permission(
    permission_id: int,
    current_user: User = Depends(get_current_user_from_token),
    db: Session = Depends(get_db),
):
    require_user_permission(
        current_user,
        "Manage Permission",
        db,
    )

    existing_permission = (
        db.query(Permission)
        .filter(Permission.id == permission_id)
        .first()
    )

    if not existing_permission:
        raise HTTPException(
            status_code=404,
            detail="Permission not found",
        )

    role_permission = (
        db.query(RolePermission)
        .filter(RolePermission.permission_id == permission_id)
        .first()
    )

    if role_permission:
        raise HTTPException(
            status_code=400,
            detail="Cannot delete permission because it is assigned to roles",
        )

    deleted_data = {
        "permission_name": existing_permission.permission_name,
        "module_name": existing_permission.module_name,
        "description": existing_permission.description,
        "status": existing_permission.status,
    }

    create_audit_log(
        db=db,
        module_name="Permission Master",
        action="Delete Permission",
        current_user=current_user,
        entity_type="Permission",
        entity_id=existing_permission.id,
        entity_label=f"{existing_permission.module_name} - {existing_permission.permission_name}",
        remarks="Permission deleted",
        request_path=f"/permissions/{permission_id}",
        details={
            "deleted": deleted_data,
        },
    )

    db.delete(existing_permission)
    db.commit()

    return {
        "message": "Permission deleted successfully"
    }


@router.post("/seed-standard")
def seed_standard_permissions(
    current_user: User = Depends(get_current_user_from_token),
    db: Session = Depends(get_db),
):
    require_user_permission(
        current_user,
        "Manage Permission",
        db,
    )

    standard_permissions = [
        # User Management
        {
            "permission_name": "View User",
            "module_name": "User Master",
            "description": "Can view users",
        },
        {
            "permission_name": "Manage User",
            "module_name": "User Master",
            "description": "Can create, update, and delete users",
        },
        {
            "permission_name": "View Role",
            "module_name": "Role Master",
            "description": "Can view roles",
        },
        {
            "permission_name": "Manage Role",
            "module_name": "Role Master",
            "description": "Can create, update, and delete roles",
        },
        {
            "permission_name": "View Permission",
            "module_name": "Permission Master",
            "description": "Can view permissions",
        },
        {
            "permission_name": "Manage Permission",
            "module_name": "Permission Master",
            "description": "Can create, update, and delete permissions",
        },
        {
            "permission_name": "View Role Permission Assignment",
            "module_name": "Role Permission Assignment",
            "description": "Can view role permission assignments",
        },
        {
            "permission_name": "Manage Role Permission Assignment",
            "module_name": "Role Permission Assignment",
            "description": "Can assign permissions to roles",
        },
        {
            "permission_name": "View User Role Assignment",
            "module_name": "User Role Assignment",
            "description": "Can view user role assignments",
        },
        {
            "permission_name": "Manage User Role Assignment",
            "module_name": "User Role Assignment",
            "description": "Can assign roles to users",
        },
        {
            "permission_name": "View Access Summary",
            "module_name": "Access Summary",
            "description": "Can view final RBAC access summary",
        },
        {
            "permission_name": "View Dashboard",
            "module_name": "Dashboard",
            "description": "Can view dashboard configurations",
        },
        {
            "permission_name": "Manage Dashboard",
            "module_name": "Dashboard",
            "description": "Can create, update, publish, and revert dashboards",
        },

        # Master Data
        {
            "permission_name": "View Location",
            "module_name": "Location Master",
            "description": "Can view locations",
        },
        {
            "permission_name": "Manage Location",
            "module_name": "Location Master",
            "description": "Can create, update, and delete locations",
        },
        {
            "permission_name": "View Location Accounting Day Setting",
            "module_name": "Location Accounting Day Setting",
            "description": "Can view location-wise accounting day settings",
        },
        {
            "permission_name": "Manage Location Accounting Day Setting",
            "module_name": "Location Accounting Day Setting",
            "description": "Can create, update, and delete location-wise accounting day settings",
        },
        {
            "permission_name": "View Asset Type",
            "module_name": "Asset Type Master",
            "description": "Can view asset types",
        },
        {
            "permission_name": "Manage Asset Type",
            "module_name": "Asset Type Master",
            "description": "Can create, update, and delete asset types",
        },
        {
            "permission_name": "View Asset",
            "module_name": "Asset Master",
            "description": "Can view assets",
        },
        {
            "permission_name": "Manage Asset",
            "module_name": "Asset Master",
            "description": "Can create, update, and delete assets",
        },
        {
            "permission_name": "View Calibration Template",
            "module_name": "Calibration Template Master",
            "description": "Can view calibration templates",
        },
        {
            "permission_name": "Manage Calibration Template",
            "module_name": "Calibration Template Master",
            "description": "Can create, update, and delete calibration templates",
        },
        {
            "permission_name": "View Asset Calibration",
            "module_name": "Asset Calibration Table",
            "description": "Can view asset calibration tables",
        },
        {
            "permission_name": "Manage Asset Calibration",
            "module_name": "Asset Calibration Table",
            "description": "Can create, upload, update, and delete calibration data",
        },
        {
            "permission_name": "View Asset Assignment",
            "module_name": "Asset Assignment",
            "description": "Can view asset assignments",
        },
        {
            "permission_name": "Manage Asset Assignment",
            "module_name": "Asset Assignment",
            "description": "Can create, update, and delete asset assignments",
        },
        {
            "permission_name": "View Asset Assignment Summary",
            "module_name": "Asset Assignment Summary",
            "description": "Can view asset assignment summary",
        },

        # Operations
        {
            "permission_name": "View Operation Type",
            "module_name": "Operations",
            "description": "Can view operation type master",
        },
        {
            "permission_name": "Manage Operation Type",
            "module_name": "Operations",
            "description": "Can create, update, and delete operation types",
        },
        {
            "permission_name": "View Tank Operation",
            "module_name": "Operations",
            "description": "Can view location-wise tank operation master",
        },
        {
            "permission_name": "Manage Tank Operation",
            "module_name": "Operations",
            "description": "Can create, update, and delete location-wise tank operations",
        },
        # Vessel / FSO soft-coded operations
        {
            "permission_name": "View Vessel Operation",
            "module_name": "Operations",
            "description": "Can view Vessel Operation Master (soft-coded Loading/Unloading/STS/Decanting etc.)",
        },
        {
            "permission_name": "Manage Vessel Operation",
            "module_name": "Operations",
            "description": "Can create, update, and delete Vessel Operation Master entries",
        },

        # Vessel Stock Ledger (Shuttle / FSO)
        {
            "permission_name": "View Vessel Stock Ledger",
            "module_name": "Operations",
            "description": "Can view Vessel Stock Ledger (approved-only derived ledger for Shuttle/FSO)",
        },

        # Movement Mapping (Barge ↔ Shuttle ↔ FSO reconciliation)
        {
            "permission_name": "View Movement Mapping",
            "module_name": "Operations",
            "description": "Can view Movement Mapping and reconciliation comparisons",
        },
        {
            "permission_name": "Manage Movement Mapping",
            "module_name": "Operations",
            "description": "Can create/update/close Movement Mapping and attach approved tickets for reconciliation",
        },
        {
            "permission_name": "View Shuttle Tracking",
            "module_name": "Operations",
            "description": "Can view Shuttle Tracking (Approved-only voyage tracking within location)",
        },
        {
            "permission_name": "Manage Shuttle Tracking",
            "module_name": "Operations",
            "description": "Can close/reopen Shuttle voyages",
        },
        {
            "permission_name": "View FSO Tracking",
            "module_name": "Operations",
            "description": "Can view FSO Tracking (Approved-only, shuttle-number based)",
        },
        {
            "permission_name": "Manage FSO Tracking",
            "module_name": "Operations",
            "description": "Can close/reopen FSO voyages",
        },
        {
            "permission_name": "View Tank Stock Ledger",
            "module_name": "Operations",
            "description": "Can view tank stock ledger and stock movement summary",
        },
        {
            "permission_name": "Manage Tank Stock Ledger",
            "module_name": "Operations",
            "description": "Can rebuild or manage tank stock ledger entries",
        },
        {
            "permission_name": "View Location Operation Availability",
            "module_name": "Operations",
            "description": "Can view location operation availability",
        },
        {
            "permission_name": "Manage Location Operation Availability",
            "module_name": "Operations",
            "description": "Can configure operation availability by location",
        },
        {
            "permission_name": "View Operation Template",
            "module_name": "Operations",
            "description": "Can view operation templates",
        },
        {
            "permission_name": "Manage Operation Template",
            "module_name": "Operations",
            "description": "Can create, update, and delete operation templates",
        },
        {
            "permission_name": "Create Operation Entry",
            "module_name": "Operations",
            "description": "Can create new operation tickets from Operation Entry",
        },
        {
            "permission_name": "View Operation Transaction",
            "module_name": "Operations",
            "description": "Can view operation transaction register and detail",
        },
        {
            "permission_name": "Submit Operation Transaction",
            "module_name": "Operations",
            "description": "Can submit draft operation tickets",
        },
        {
            "permission_name": "Review Operation Transaction",
            "module_name": "Operations",
            "description": "Can review operation tickets before submit/approve confirmation",
        },
        {
            "permission_name": "Approve Operation Transaction",
            "module_name": "Operations",
            "description": "Can approve submitted operation tickets",
        },
        {
            "permission_name": "Reject Operation Transaction",
            "module_name": "Operations",
            "description": "Can reject submitted operation tickets",
        },
        {
            "permission_name": "Cancel Operation Transaction",
            "module_name": "Operations",
            "description": "Can cancel draft or rejected operation tickets",
        },
        {
            "permission_name": "Request Approved Transaction Correction",
            "module_name": "Operations",
            "description": "Can request admin revoke for approved operation tickets that need correction",
        },
        {
            "permission_name": "View Approved Transaction Correction Requests",
            "module_name": "Operations",
            "description": "Can view approved transaction correction and revoke requests",
        },
        {
            "permission_name": "Admin Revoke Approved Transaction",
            "module_name": "Operations",
            "description": "Can revoke approval and push approved tickets back to submitted review",
        },
        # Barge Seal Master
        {
            "permission_name": "View Barge Seal Master",
            "module_name": "Barge Seal Master",
            "description": "Can view barge seal master",
        },
        {
            "permission_name": "Manage Barge Seal Master",
            "module_name": "Barge Seal Master",
            "description": "Can create/update barge seal master",
        },
        # Flowmeter Configuration / Records
        {
            "permission_name": "View Flowmeter Config",
            "module_name": "Flowmeter Config",
            "description": "Can view flowmeter meter configuration by location and asset",
        },
        {
            "permission_name": "Manage Flowmeter Config",
            "module_name": "Flowmeter Config",
            "description": "Can create, update, and delete flowmeter meter configuration",
        },
        {
            "permission_name": "View Flowmeter Record",
            "module_name": "Flowmeter Record",
            "description": "Can view flowmeter meter records",
        },
        {
            "permission_name": "Create Flowmeter Record",
            "module_name": "Flowmeter Record",
            "description": "Can create flowmeter meter records",
        },
        # Company / Report Profiles
        {
            "permission_name": "View Company Report Profile",
            "module_name": "Company Report Profile",
            "description": "Can view company report profiles used for printable reports",
        },
        {
            "permission_name": "Manage Company Report Profile",
            "module_name": "Company Report Profile",
            "description": "Can create, update, and delete company report profiles",
        },

        # Audit Logs
        {
            "permission_name": "View Audit Log",
            "module_name": "Audit Log",
            "description": "Can view system audit logs",
        },

        # System Notifications
        {
            "permission_name": "View System Notification",
            "module_name": "System Notification",
            "description": "Can view system notifications and user notification inbox",
        },
        {
            "permission_name": "Manage System Notification",
            "module_name": "System Notification",
            "description": "Can create and update system notifications",
        },
        {
            "permission_name": "Publish System Notification",
            "module_name": "System Notification",
            "description": "Can publish system notifications to users",
        },
        {
            "permission_name": "Deactivate System Notification",
            "module_name": "System Notification",
            "description": "Can deactivate published system notifications",
        },
        {
            "permission_name": "Acknowledge System Notification",
            "module_name": "System Notification",
            "description": "Can acknowledge or dismiss assigned system notifications",
        },
        {
            "permission_name": "View System Notification Delivery Report",
            "module_name": "System Notification",
            "description": "Can view delivery, seen, dismissed, and acknowledgement reports",
        },

        # Backup Recovery
        {
            "permission_name": "View Backup",
            "module_name": "Backup Recovery",
            "description": "Can view backup settings and backup job history",
        },
        {
            "permission_name": "Create Manual Backup",
            "module_name": "Backup Recovery",
            "description": "Can create a manual database backup",
        },
        {
            "permission_name": "Manage Backup Settings",
            "module_name": "Backup Recovery",
            "description": "Can configure automatic backup schedule and retention settings",
        },
        {
            "permission_name": "Download Backup",
            "module_name": "Backup Recovery",
            "description": "Can download completed backup files",
        },
        {
            "permission_name": "Verify Backup Checksum",
            "module_name": "Backup Recovery",
            "description": "Can verify backup file checksum integrity",
        },
        {
            "permission_name": "Delete Backup",
            "module_name": "Backup Recovery",
            "description": "Can delete backup files while preserving backup job history",
        },
        {
            "permission_name": "Run Backup Cleanup",
            "module_name": "Backup Recovery",
            "description": "Can run retention cleanup for old backup files",
        },
        {
            "permission_name": "Request Backup Restore",
            "module_name": "Backup Recovery",
            "description": "Can request restore approval for a completed backup",
        },
        {
            "permission_name": "Approve Backup Restore",
            "module_name": "Backup Recovery",
            "description": "Can approve backup restore requests without executing restore",
        },
        {
            "permission_name": "Reject Backup Restore",
            "module_name": "Backup Recovery",
            "description": "Can reject backup restore requests",
        },
        {
            "permission_name": "Validate Backup Restore",
            "module_name": "Backup Recovery",
            "description": "Can validate an approved backup restore request against a separate validation database",
        },
        {
            "permission_name": "Execute Backup Restore",
            "module_name": "Backup Recovery",
            "description": "Can execute a validated backup restore into the production database",
        },

        # Reports / Admin Future
        {
            "permission_name": "View Reports",
            "module_name": "Reports",
            "description": "Can view reports",
        },
        {
            "permission_name": "Export Reports",
            "module_name": "Reports",
            "description": "Can export reports",
        },
        {
            "permission_name": "View Admin Settings",
            "module_name": "Admin",
            "description": "Can view admin settings",
        },
        {
            "permission_name": "Manage Admin Settings",
            "module_name": "Admin",
            "description": "Can manage admin settings",
        },
        {
            "permission_name": "View Out-Turn Report",
            "module_name": "Reports",
            "description": "Can view Out-Turn Report from approved tank stock ledger rows",
        },
        {
            "permission_name": "View Material Balance Report",
            "module_name": "Reports",
            "description": "Can view Material Balance Report from tank stock ledger",
        },
        {
            "permission_name": "View Material Balance Template",
            "module_name": "Configuration",
            "description": "Can view Material Balance template configuration",
        },
        {
            "permission_name": "Manage Material Balance Template",
            "module_name": "Configuration",
            "description": "Can create, edit, and delete Material Balance template configuration",
        },
        {
            "permission_name": "View Operation Workflow Policy",
            "module_name": "Operations",
            "description": "Can view operation workflow authorization policies",
        },
        {
            "permission_name": "Manage Operation Workflow Policy",
            "module_name": "Operations",
            "description": "Can create, update, and delete operation workflow authorization policies",
        },
        {
            "permission_name": "View My Tasks",
            "module_name": "Operations",
            "description": "Can view assigned operation approval tasks",
        },
        {
            "permission_name": "Act On Operation Task",
            "module_name": "Operations",
            "description": "Can take, release, approve, or reject assigned operation tasks",
        },
        {
            "permission_name": "Manage Operation Tasks",
            "module_name": "Operations",
            "description": "Can view and manage all operation approval tasks",
        },
        {
            "permission_name": "View Own Security Settings",
            "module_name": "User Security",
            "description": "Can view own password and 2FA security settings",
        },
        {
            "permission_name": "Manage Own Security Settings",
            "module_name": "User Security",
            "description": "Can change own password and manage own 2FA settings",
        },
        {
            "permission_name": "Request Password Reset",
            "module_name": "User Security",
            "description": "Can request administrator password reset",
        },
        {
            "permission_name": "Reset User Password",
            "module_name": "User Security",
            "description": "Can hard reset user passwords",
        },
        {
            "permission_name": "Reset User 2FA",
            "module_name": "User Security",
            "description": "Can reset user 2FA during administrator password reset",
        },
    ]

    created_count = 0
    existing_count = 0

    for permission_data in standard_permissions:
        existing_permission = (
            db.query(Permission)
            .filter(
                Permission.permission_name.ilike(permission_data["permission_name"]),
                Permission.module_name.ilike(permission_data["module_name"]),
            )
            .first()
        )

        if existing_permission:
            existing_count += 1
            continue

        new_permission = Permission(
            permission_name=permission_data["permission_name"],
            module_name=permission_data["module_name"],
            description=permission_data["description"],
            status="Active",
        )

        db.add(new_permission)
        created_count += 1

    create_audit_log(
        db=db,
        module_name="Permission Master",
        action="Seed Standard Permissions",
        current_user=current_user,
        entity_type="Permission",
        entity_id=None,
        entity_label="Standard Permission Seed",
        remarks="Seeded standard permissions",
        request_path="/permissions/seed-standard",
        details={
            "created_count": created_count,
            "existing_count": existing_count,
            "total_standard_permissions": len(standard_permissions),
        },
    )

    db.commit()

    return {
        "message": "Standard permissions seed completed",
        "created_count": created_count,
        "existing_count": existing_count,
        "total_standard_permissions": len(standard_permissions),
    }