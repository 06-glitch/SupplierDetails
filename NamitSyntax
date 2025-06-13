string ls_err_msg
string ls_storage_loc_nbr
string ls_location_exists
string ls_current_storage_loc_type_code
string ls_active_status_ind
string ls_storage_loc_type_code
long li_return
long ll_rc

ls_storage_loc_nbr = sle_storage_loc_nbr.text

IF ls_storage_loc_nbr = '' THEN
    messagebox("Usage Error", "Trailer Nbr is a required field.")
    RETURN
END IF

IF rb_transit.Checked THEN
    ls_storage_loc_type_code = "TRANSIT"
ELSEIF rb_storage.Checked THEN
    ls_storage_loc_type_code = "TRAILBULK"
ELSE
    messagebox("Usage Error", "Trailer Type is required.")
    RETURN
END IF

ls_location_exists = "N"
SELECT 'Y', Storage_Loc_Type_Code, Active_Status_Ind
INTO :ls_location_exists, :ls_current_storage_loc_type_code, :ls_active_status_ind
FROM t_Storage_Location
WHERE t_Storage_Location.Storage_Loc_Nbr = :ls_storage_loc_nbr;

IF SQLCA.sqlcode < 0 THEN
    ls_err_msg = sqlca.sqlerrtext
    messagebox("Error", "Looking for Storage Location information: " + ls_err_msg)
    RETURN
END IF

IF (ls_location_exists = "Y") THEN
    IF (ls_current_storage_loc_type_code <> "TRAILBULK") AND (ls_current_storage_loc_type_code <> "TRANSIT") THEN
        messagebox("Usage Error", "Storage Location " + ls_storage_loc_nbr + " already exists and is not a trailer type location.")
        RETURN
    END IF

    IF (ls_active_status_ind = "I") THEN
        ll_rc = MessageBox("Warning", "In-Active Storage Location for Trailer Nbr " + ls_storage_loc_nbr + " already exists.~r~nWould you like to re-activate it?", Exclamation!, YesNo!, 1)
        IF ll_rc = 1 THEN
            UPDATE t_Storage_Location
            SET Active_Status_Ind = 'A',
                Storage_Loc_Type_Code = :ls_storage_loc_type_code
            WHERE t_Storage_Location.Storage_Loc_Nbr = :ls_storage_loc_nbr
            USING SQLCA;

            IF SQLCA.sqlcode <> 0 THEN
                ls_err_msg = sqlca.sqlerrtext
                messagebox("Error", "Update of Storage Location~''s Active_Status_Ind failed: " + ls_err_msg)
                RETURN
            END IF
        ELSE
            RETURN // canceled
        END IF

    ELSEIF (ls_current_storage_loc_type_code <> ls_storage_loc_type_code) THEN
        ll_rc = MessageBox("Warning", "Are you sure you want to change the Storage Location Type Code from " + Trim(ls_current_storage_loc_type_code) + " to " + ls_storage_loc_type_code + "?", Exclamation!, YesNo!, 1)
        IF ll_rc = 1 THEN
            UPDATE t_Storage_Location
            SET Storage_Loc_Type_Code = :ls_storage_loc_type_code
            WHERE t_Storage_Location.Storage_Loc_Nbr = :ls_storage_loc_nbr
            USING SQLCA;

            IF SQLCA.sqlcode <> 0 THEN
                ls_err_msg = sqlca.sqlerrtext
                messagebox("Error", "Update of Storage Location~''s Storage_Loc_Type_Code failed: " + ls_err_msg)
                RETURN
            END IF
        ELSE
            RETURN // canceled
        END IF
    END IF

ELSE
    SQLCA.ap_add_trailer_storage_loc(ls_storage_loc_nbr, ls_storage_loc_type_code)
    IF SQLCA.sqlcode <> 0 THEN
        ls_err_msg = sqlca.sqlerrtext  
        messagebox("Usage Error", "ap_Add_Trailer_Storage_Loc failed: " + ls_err_msg)
        RETURN
    END IF
END IF

// Update the Warehouse Storage Handling Codes
IF ls_storage_loc_type_code = "TRANSIT" AND Left(ls_storage_loc_nbr, 2) = "OB" THEN
    DELETE FROM t_Storage_Loc_Whse_Storage_Handling_Xref
    WHERE Storage_Loc_Nbr = :ls_storage_loc_nbr;

    IF SQLCA.sqlcode <> 0 THEN
        ls_err_msg = sqlca.sqlerrtext
        messagebox("Usage Warning", "Cleanup of Warehouse Storage Handling Codes failed. Contact Inventory Accounting: " + ls_err_msg)
    END IF

    INSERT INTO t_Storage_Loc_Whse_Storage_Handling_Xref (MMM_Facility_Code, Storage_Loc_Nbr, Warehouse_Storage_Handling_Code)
        SELECT 'HT', :ls_storage_loc_nbr, '2'
        UNION SELECT 'HT', :ls_storage_loc_nbr, '9'
        UNION SELECT 'HT', :ls_storage_loc_nbr, 'A'
        UNION SELECT 'HT', :ls_storage_loc_nbr, 'E'
        UNION SELECT 'HT', :ls_storage_loc_nbr, 'F';

    IF SQLCA.sqlcode <> 0 THEN
        ls_err_msg = sqlca.sqlerrtext
        messagebox("Usage Warning", "Insert of Warehouse Storage Handling Codes failed. Contact Inventory Accounting: " + ls_err_msg)
    END IF

ELSEIF ls_storage_loc_type_code = "TRAILBULK" THEN
    DELETE FROM t_Storage_Loc_Whse_Storage_Handling_Xref
    WHERE Storage_Loc_Nbr = :ls_storage_loc_nbr;

    IF SQLCA.sqlcode <> 0 THEN
        ls_err_msg = sqlca.sqlerrtext
        messagebox("Usage Warning", "Cleanup of Warehouse Storage Handling Codes failed. Contact Inventory Accounting: " + ls_err_msg)
    END IF

    INSERT INTO t_Storage_Loc_Whse_Storage_Handling_Xref (MMM_Facility_Code, Storage_Loc_Nbr, Warehouse_Storage_Handling_Code)
        SELECT 'HT', :ls_storage_loc_nbr, 'A'
        UNION SELECT 'HT', :ls_storage_loc_nbr, 'E';

    IF SQLCA.sqlcode <> 0 THEN
        ls_err_msg = sqlca.sqlerrtext
        messagebox("Usage Warning", "Insert of Warehouse Storage Handling Codes failed. Contact Inventory Accounting: " + ls_err_msg)
    END IF
END IF

IF Parent.cbx_print_barcode.Checked THEN
    Parent.triggerEvent("ue_barcode_rpt")
END IF

close(parent)
RETURN
