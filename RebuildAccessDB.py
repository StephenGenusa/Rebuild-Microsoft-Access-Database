import argparse
import os
import shutil
import subprocess
import sys
import time
import urllib.request
import win32con
import win32gui
import win32process

import win32com.client as win32

# AcCloseSave enumeration
acSavePrompt = 0
acSaveYes = 1
acSaveNo = 2

# AcObjectType enumeration
acTable = 0
acQuery = 1
acForm = 2
acReport = 3
acMacro = 4
acModule = 5

# AcTextTransferType enumeration
acExportDelim = 2


def update_vbac_script(vbac_full_filename):
    """Modify script, if necessary, to allow queries to be exported as well"""
    QUERY_PARAM = "param.incQuery = false;"
    with open(vbac_full_filename, mode='r') as script_file:
        script_text = script_file.read()
    if "//" + QUERY_PARAM not in script_text:
        script_text = script_text.replace(QUERY_PARAM, "//" + QUERY_PARAM)
        with open(vbac_full_filename, mode='w') as script_file:
            script_file.write(script_text)


def close_ms_access():
    """PostMessage to Access to exit to close it down before we try COM API"""
    attempts_to_close = 0
    while win32gui.FindWindow("OMain", None) and attempts_to_close < 15:
        win32gui.PostMessage(win32gui.FindWindow("OMain", None), win32con.WM_CLOSE, 0, 0)
        time.sleep(1e-3)  # Yeah, it's a bit hacky
        attempts_to_close += 1


def open_ms_access(ms_access_filename, window_visible):
    """Create an instance of Access.Application"""
    ms_access = win32.dynamic.Dispatch("Access.Application")
    if ms_access_filename:
        ms_access.OpenCurrentDatabase(ms_access_filename)
    ms_access.Visible = window_visible
    return ms_access


def save_table_data(ms_access_filename, export_path):
    """Exports data from all tables to export_path"""
    ms_access = open_ms_access(ms_access_filename, True)
    if ms_access.CurrentDb.TableDefs.Count > 0:
        for table_def in ms_access.CurrentDb.TableDefs:
            if table_def.Name[:5] != "MSys":
                ms_access.DoCmd.TransferText(acExportDelim, None, table_def.Name,
                                             os.path.join(export_path, "TableData", "Table_" + table_def.Name & ".txt"),
                                             True)


# def save_query_defs(ms_access_filename, export_path):
#     """Exports query definitions to export_path. Found support for this in vbac.wsf in digging through the file"""
#     ms_access = open_ms_access(ms_access_filename, True)
#     print("Saving Query Definitions")
#     if ms_access.CurrentData.AllQueries.Count > 0:
#         for query_def in ms_access.CurrentData.AllQueries:
#             export_filename = os.path.join(export_path, query_def.Name + ".txt")
#             print("-", query_def.Name)
#             ms_access.SaveAsText(acQuery, query_def.Name, export_filename)
#     ms_access.Quit(acSaveNo)


def import_query_defs(ms_access_filename, import_path):
    """Import the query definitions from import_path"""
    ms_access = open_ms_access(ms_access_filename, True)
    file_list = [os.path.join(import_path, x) for x in os.listdir(import_path) if ".qry" in x.lower()]
    if len(file_list) > 0:
        print("Import Query Definitions")
        for file_name in file_list:
            query_name = os.path.splitext(os.path.basename(file_name))[0]
            print('- Import:', os.path.basename(file_name))
            ms_access.LoadFromText(acQuery, query_name, file_name)
    else:
        print("** WARNING: No Query Definitions Found **")
    ms_access.Quit(acSaveYes)


def get_vba_references(ms_access_filename):
    """Get VBA references from original database"""
    ms_access = open_ms_access(ms_access_filename, True)
    ref_paths = []
    if ms_access.References.Count > 0:
        for reference in ms_access.References:
            ref_paths.append(reference.FullPath)
    ms_access.Quit(acSaveNo)
    return ref_paths


def reference_already_exists(access_references, new_reference):
    """Test if a given VBA reference already exists in the new database. Only needed if db created blank"""
    for reference in access_references:
        if reference.FullPath.lower() == new_reference.lower():
            return True
    return False


def set_vba_references(ms_access_filename, refs_to_include):
    """Set the VBA references"""
    ms_access = open_ms_access(ms_access_filename, True)
    for ref_path in refs_to_include:
        if not reference_already_exists(ms_access.References, ref_path):
            print(ref_path)
            ms_access.References.AddFromFile(ref_path)
    ms_access.Quit(acSaveYes)


def access_compact_and_repair(ms_access_filename):
    """Call the Microsoft Access compact and repair function"""
    dest_db_filename = ms_access_filename.replace(".accdb", "_backup.accdb")
    ms_access = open_ms_access("", False)
    ms_access.compactRepair(ms_access_filename, dest_db_filename)
    os.remove(ms_access_filename)
    os.rename(dest_db_filename, ms_access_filename)
    ms_access.Quit()


def delete_old_access_objects(ms_access_filename):
    """Deletes existing forms, reports, modules and queries. Retains table definitions"""
    def delete_documents_from_container(access_object_type, container, container_name):
        # COM has a problem with indexes deleting forward so reverse the elements and work
        # bottom to top. This doesn't have the same problem...
        print(container_name)
        for document_object in reversed(container):
            try:
                print("- Deleting:", document_object.Name)
                ms_access.DoCmd.Close(access_object_type, document_object.Name, acSaveNo)
            except Exception:
                pass
            try:
                ms_access.DoCmd.DeleteObject(access_object_type, document_object.Name)
            except Exception:
                pass

    ms_access = win32.dynamic.Dispatch("Access.Application")
    ms_access = open_ms_access(ms_access_filename, True)
    delete_documents_from_container(acForm, ms_access.CurrentProject.AllForms, "Forms")
    delete_documents_from_container(acReport, ms_access.CurrentProject.AllReports, "Reports")
    delete_documents_from_container(acModule, ms_access.CurrentProject.AllModules, "Modules")
    delete_documents_from_container(acQuery, ms_access.CurrentData.AllQueries, "Queries")
    ms_access.Quit(acSaveYes)
    ms_access = None


def decombine_microsoft_access(source_accdb_filename):
    """Calls vbac.wsf to convert objects into text files"""
    subprocess.call(r'cscript vbac.wsf decombine /incQuery:true /binary:"' + source_accdb_filename + '"')


def import_ms_access_assets():
    """Imports forms, reports and modules. Queries handled in this Python module"""
    subprocess.call(r"cscript vbac.wsf combine")


def verify_directory_exists(directory_path):
    if not os.path.exists(directory_path):
        print("Setting up directory", directory_path)
        os.makedirs(directory_path)


def delete_all_files_except(filename_to_save):
    path_to_check = os.path.dirname(filename_to_save)
    file_list = [os.path.join(path_to_check, x) for x in os.listdir(path_to_check)]
    if len(file_list) > 1:
        print("Removing files with no data definitions")
    for file_name in file_list:
        if file_name != filename_to_save:
            os.remove(file_name)


def main():
    close_ms_access()
    parser = argparse.ArgumentParser(prog=os.path.splitext(os.path.basename(__file__))[0],
                                     usage='%(prog)s [options] path',
                                     description='Rebuild Microsoft Access Database')
    parser.add_argument('-i', '--input-file', action='store', type=str, required=True, dest="input_file")
    parser.add_argument('-c', '--create-new-db', action='store', type=bool, default=False, dest="create_new_db")
    parser.add_argument('-d', '--download-script', action='store', type=bool, default=False, dest="download_script")
    parser.add_argument('-v', "--version", action='version', version='%(prog)s 1.0')
    args = parser.parse_args()

    if not os.path.isfile(args.input_file):
        parser.print_help()
        print("Error. File does not exist:", args.input_file)
        sys.exit()

    # Setup directories
    ariawase_path = os.path.join(os.environ.get("USERPROFILE"), "ariawase")
    dest_accdb_filename = os.path.join(ariawase_path, "bin", os.path.basename(args.input_file))

    ariawase_src_path = os.path.join(ariawase_path, "src")
    shutil.rmtree(ariawase_src_path)
    verify_directory_exists(os.path.join(ariawase_path, "bin"))
    verify_directory_exists(ariawase_src_path)

    # Make sure vbac.wsf is available
    if not os.path.isfile(os.path.join(ariawase_path, "vbac.wsf")):
        if args.download_script:
            print("Grabbing vbac.wsf from github")
            url = "https://raw.githubusercontent.com/vbaidiot/ariawase/master/"
            urllib.request.urlretrieve(url + "vbac.wsf", filename=os.path.join(ariawase_path, "vbac.wsf"))
            urllib.request.urlretrieve(url + "LICENSE.txt", filename=os.path.join(ariawase_path, "LICENSE.txt"))
        else:
            print("ariawase vbac.wsf is not available. Please download or use -d option")
            sys.exit()

    # Set directory for vbac.wsf execution
    os.chdir(ariawase_path)

    # Note: ariawase can either create a new empty accdb or use one that is already in the bin directory
    if not args.create_new_db:
        # Copy the source database to the ariawase\bin directory
        shutil.copyfile(args.input_file, dest_accdb_filename)
        access_compact_and_repair(dest_accdb_filename)
        # Remove all forms, reports and modules
        delete_old_access_objects(dest_accdb_filename)
        # Compact and Repair...
        for _ in range(0, 3):
            access_compact_and_repair(dest_accdb_filename)

    # Convert original forms, reports, modules and queries to text files
    update_vbac_script(os.path.join(ariawase_path, "vbac.wsf"))
    decombine_microsoft_access(os.path.dirname(args.input_file))

    # Convert original queries to text files
    # save_query_defs(args.input_file, query_src_path)

    # Import queries
    import_query_defs(dest_accdb_filename, os.path.join(ariawase_src_path, os.path.basename(dest_accdb_filename)))

    # Import form, report and module text files into Access
    import_ms_access_assets()

    if args.create_new_db:
        # Set VBA references and rebuild tabledefs and queries
        # - Get the VBA references in the original accdb
        refs_to_restore = get_vba_references(args.input_file)

        # - Set the VBA references on the new db
        set_vba_references(os.path.join(ariawase_path, "bin", os.path.basename(args.input_file)), refs_to_restore)

        # Add data - NOT IMPLEMENTED YET
        # https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/tabledef-object-dao
    else:
        # Remove any empty database def files created by ariawase
        delete_all_files_except(dest_accdb_filename)

    # Final compact and repairs
    for _ in range(0, 3):
        access_compact_and_repair(dest_accdb_filename)
        
    print("Rebuild completed")
    subprocess.Popen(r'explorer /select,' + dest_accdb_filename)

if __name__ == '__main__':
    main()
