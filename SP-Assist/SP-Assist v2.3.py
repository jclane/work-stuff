import sys, os, time, csv, threading
from multiprocessing.dummy import Pool as ThreadPool

BRANDS = {1:('ACE', 'GWY', 'EMA'), 2:'APP', 3:'ASU',4:('DEL','DRC'),5:('HEW','COM'),6:'LEN',7:'SAC',8:'SYC',9:'TSC', 10:('ACE', 'GWY', 'EMA', 'APP', 'ASU','DEL','DRC','HEW','COM','LEN','SAC','SYC','TSC')}
SEARCH_MENU = {0:'\n#### STEP 2 | What would you like to search? ####\n',1:'\t1. Part Classes', 2:'\t2. Part Numbers', 3:'\t3. Search Descriptions', 4:'\t4. Search All', 5:'\t5. Quit\n'}      
MAIN_MENU = {0:'\n\nPlease make a selection:\n', 1:'\t1. Search files', 2:'\t2. Seach and replace', 3:'\t3. Generate bad class list (coming soon)', 4:'\t4. Quit\n'}
BRAND_MENU = {0:'\n#### STEP 1 | Which brand would you like search? ####\n', 1:'\t1. Acer/Gateway/Emachine', 2:'\t2. Apple', 3:'\t3. Asus', \
            4:'\t4. Dell/Dell Reclamation', 5:'\t5. HP/Compaq', 6:'\t6. Lenovo', 7:'\t7. Samsung', 8:'\t8. Sony', 9:'\t9. Toshiba', \
            10:'\t10. Canandian BOMs', 11:'\t11. All TEH BRANDS!!1!', 12:'\t12. Quit\n'}
            
# Begin often used functions       
def clear_screen():
    os.system('cls')

def run_time(start_time):
    return round(time.time() - start_time)
    
def print_menu(menu_to_print):  
    
    for el in menu_to_print:
        print(menu_to_print[el])
    
    while True:
        try:               
            selection = int(input('Select Number (1-' + str(len(menu_to_print) - 1) + '): '))
            if not selection > len(menu_to_print) - 1 and not selection < 1:
                break
            else:
                raise ValueError
        except ValueError:
            print('\n\n!ERROR : Please select a number from the menu.')
            continue       
    
    if selection == len(menu_to_print) - 1:
        print('\nQuitting')
        raise SystemExit

    return selection
    
def class_is_valid(partclass):
    CLASSES = ['AC', 'AD', 'ANT', 'AUBD', 'AUDIO', 'BAT', 'BLUE', 'BRA', 'BRD', 'CAB', 'CAM', 
    'CARD', 'CD', 'CMOS', 'CORD', 'COS', 'CRBD', 'DCBD', 'DCJK', 'DOCKIT', 'DVD', 'FAN', 'FDD', 
    'FLD', 'GK', 'HDD', 'HEAT SYNC', 'INV', 'IOB', 'IOP', 'KB', 'LAN', 'LCD', 'LDBD', 'LK', 'MEM', 
    'MIC', 'MOD', 'OBRD', 'OTHER', 'PEN', 'PROC', 'PWBD', 'PWR', 'REM', 'SCRD', 'TABMB', 'TUBRD', 
    'USBD', 'VBRD', 'VGBD', 'WIR'] 
               
    if partclass not in CLASSES:
        return False
    else:
        return True
# End often used functions

def get_files(brand):
    """ Set 'rootdir' and return list of files to be searched in variable 'fileslist'. """
    fileslist = []
    
    if brand == 1:
        rootdir = r'!!REDACTED!!\Acer_Gateway_Emachine'
    elif brand == 2:
        rootdir = r'!!REDACTED!!\Apple-APP'
    elif brand == 3:
        rootdir = r'!!REDACTED!!\Asus-ASU'
    elif brand == 4:
        rootdir = r'!!REDACTED!!\Dell_Dell Reclamation'
    elif brand == 5:
        rootdir = r'!!REDACTED!!\HP_Compaq'
    elif brand == 6:
        rootdir = r'!!REDACTED!!\Lenovo-LEN'
    elif brand == 7:
        rootdir = r'!!REDACTED!!\Samsung Computer-SAC'
    elif brand == 8:
        rootdir = r'!!REDACTED!!\Sony Computer-SYC'
    elif brand == 9:
        rootdir = r'!!REDACTED!!\Toshiba Computer-TSC'
    elif brand == 10:
        rootdir = r'!!REDACTED!!\Canada\Computers'
    elif brand == 11:
        rootdir = r'!!REDACTED!!'
    else:
        print('ERROR')
        
    for subdir, dirs, files in os.walk(rootdir):          
        for file in files:
            fileSuffix = file.split("-")
            fileSuffix = str(fileSuffix[-1][:3])
            if brand == 11: 
                if file.endswith(".csv") and fileSuffix in BRANDS.values():
                    fileslist.append(subdir + os.sep + file)
            else:
                if file.endswith(".csv") and fileSuffix in BRANDS[brand]:     
                    fileslist.append(subdir + os.sep + file)
            
    return fileslist
     
def main(search_type, brand):
    lock = threading.Lock()
    start_time = time.time()
    search_what = print_menu(SEARCH_MENU)
    search_term = input('\nSearch for: ')
    fileslist = get_files(brand)
    found = []
    edited = []

    def csv_writer(file, rows):
        lock.acquire()
        try:
            with open(file, 'w', newline='', encoding='latin1') as csvfile:
                writer = csv.writer(csvfile)
                
                if os.path.basename(file) == 'search results.csv':
                    if len(rows[0]) == 5:
                        writer.writerow(["File Path", "Search Term", "Part Class", "Part Number", "Part Description"])
                    else:
                        writer.writerow(["File Name", "Change", "Old Line", "New Line"])

                writer.writerows(rows)
        finally:
            lock.release()

    def show_results(found, edited, total_runtime):
        '''Displays results of search indicating the number of files 
        found that match the search_term.  Also displays the number
        of rows edited and the total runtime.'''

        if len(found) > 0:
            clear_screen()
            print('\nTOTAL RUN TIME: ' + str(total_runtime) + ' seconds')
            print('\nMatches were found in ' + str(len(found)) + ' files.')
            
            if len(edited) > 0:
                print(str(len(edited)) + ' lines were edited.')
            ask = input('\nSave to results to file? (y/n) : ')
            
            if ask.lower() == 'y' or ask.lower() == 'yes':
                savedir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "search results.csv")
                csv_writer(savedir, found)
                if len(edited) > 0: csv_writer(savedir, edited)
                print('Results saved to ' + savedir)
            else:
                ask = input('Display results in command prompt? (y/n) : ')
                print('\n')
                if ask.lower() == 'y' or ask.lower() == 'yes':
                    
                    for item in found:
                        file_name = item[0]
                        #search_term = item[1]
                        #part_class = item[2]
                        #part_num = item[3]
                        #part_desc = item[4]
                        print("{} found in file {}".format(search_term, file_name))
                    print('\n')
                    if len(edited) > 0: 
                        for item in edited:                        
                            print('In file ' + item[0], item[1])
                            print('\tOld line: ' + item[2])
                            print('\tNew line: ' + item[3])
                            print('\n')
                    print('\n')
                else:
                    response = input('Return to main menu? : ')
                    if response.lower() == 'y' or response.lower() == 'yes':
                        main_menu()
                    else:
                        raise SystemExit
        else:
            print('\nTOTAL RUN TIME: ' + str(total_runtime) + ' seconds')
            print('No matches found.')
            response = input('Return to the Main Menu? (y/n) : ')
            if response.lower() == 'y' or response.lower() == 'yes':
                main_menu()
            else:
                raise SystemExit     
  
    def search_files(file): # need to figure out how to add "bad class list" option
        '''Searches the files in fileslist for matches to search_term
        and then stores the path, search_term, part class, part number,
        and description in the found[] list.'''
        
        def edit_row(row):
            lock.acquire()
            try:
                parts_of_row = {0:"class", 1:"part number", 2:"part description", 3:"entire line"}
                old_row = row.copy()
                
                clear_screen()
                print('MATCH FOUND!')
                
                print('\nDirectory: {}'.format(os.path.dirname(os.path.abspath(file))))
                print('Filename: {}'.format(os.path.basename(file)))
                
                print('\nPart Class: {}\nPart Number: {}\nDescription: {}'.format(old_row[0], old_row[1], old_row[2]))
                
                response = input("\nWould you like to edit the line? (y/n) : ")
                if response.lower() == 'y' or response.lower() == 'yes':    
                    print("-" * 41)
                    print("\nWhich part would you like to change?\n")
                    print("\n\t1. Part Class\n\t2. Part Number\n\t3. Part Description\n\t4. Whole Line\n")                    
                    which_part = int(input('(1-4): '))
                    while which_part not in range(1,5):
                        which_part = int(input("Please enter a valid selection (1-4): "))
                    
                    which_part = which_part - 1 
                    
                    new_str = input('Change to : ')
                    if which_part == 0:
                        while not class_is_valid(new_str):
                            new_str = input('\nEnter a valid part class : ')
                            
                    row[which_part] = new_str
                    edited.append([str(os.path.basename(file)), old_row[which_part] + ' was changed to ' + new_str, ','.join(old_row), ','.join(row)])

                rows.append(row)
            finally:
                lock.release()
        
        def fix_nulls(s, *kwargs):
            for line in s:
                yield line.replace('\0', '')
                                
        with open(file, "r", encoding="latin1") as csvfile:
            reader = csv.reader(fix_nulls(csvfile), quotechar='"', skipinitialspace=True)
            rows = []   
            
            try:
                for row in reader:
                
                    if row and len(row) > 2:
                        part_class = row[0]
                        part_num = row[1]
                        part_desc = row[2]
                                                    
                        if search_what == 1:
                            if search_term == part_class:
                                if search_type == 2: edit_row(row)
                                found.append([str(os.path.basename(file)), search_term, part_class, part_num, part_desc])
                            elif search_type == 2: rows.append(row)
                        elif search_what == 2:
                            if search_term == part_num:
                                if search_type == 2: edit_row(row)
                                found.append([str(os.path.basename(file)), search_term, part_class, part_num, part_desc])
                            elif search_type == 2: rows.append(row)
                        elif search_what == 3:
                            if search_term.lower() in part_desc.lower():
                                if search_type == 2: edit_row(row)
                                found.append([str(os.path.basename(file)), search_term, part_class, part_num, part_desc])
                            elif search_type == 2: rows.append(row)
                        elif search_what == 4: # part line
                            if search_term in row:
                                if search_type == 2: edit_row(row)
                                found.append([str(os.path.basename(file)), search_term, part_class, part_num, part_desc])
                            elif search_type == 2: rows.append(row)
                        else: 
                            continue
                    else:
                        continue
                      
            except TypeError as e:
                print("Error: " + str(e))
                print("@ file " + str(os.path.basename(file)))
                print("& row " + str(row)) 
            
            if search_type == 2: csv_writer(file, rows)
    
    pool = ThreadPool(4)
    pool.map(search_files, fileslist)
    pool.close()
    pool.join()

    found = [item for item in found if item is not None]
    total_runtime = run_time(start_time)
    show_results(found, edited, total_runtime)
    
def main_menu():
    clear_screen()
    search_type = print_menu(MAIN_MENU)
    clear_screen()
    brand = print_menu(BRAND_MENU)
    clear_screen()
    main(search_type, brand)

main_menu()
