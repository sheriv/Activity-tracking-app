import sys
reload(sys)  
sys.setdefaultencoding('CP1251')
import win32gui
import time
from mx import DateTime
import xlsxwriter


storage = {}
date = DateTime.now().date
print_counter = 0
print_counter_threshold = 150
txtfile = date + '.txt'
xlsxfile = date + '.xlsx'


def reset(storage, name):

    for key in storage:
      
        if key != name:
           
            if storage[key][3] == True:
            
                if storage[key][4] == 'count':

                    storage[key][3] = False
                
                else:

                    storage[key][4] += storage[key][2]
                    storage[key][3] = False
            
            else:

                pass
                  
            
        else:

            pass

def print_dict(storage, txtfile, xlsxfile):

    sortedk = sorted(storage, key = lambda k: storage[k][4], reverse = True)

    try:
        
        workbook = xlsxwriter.Workbook(xlsxfile)
        worksheet = workbook.add_worksheet()
        row = 0
        col = 0
        i = 0
        
        for key in sortedk:
            
            if storage[key][4] != 'count' and key != '':
                string2 = key
                worksheet.write(row, col, key)
                worksheet.write(row, col + 1, str(storage[key][4]))
                row += 1
            
            else:
                
                pass
                
        workbook.close()
        
        return 1
        
    except IOError:
     
        return -1
        
while True:
 
    name = win32gui.GetWindowText (win32gui.GetForegroundWindow())
 
    if name in storage:
        
        reset(storage, name)
        
        if storage[name][3] == True:
        
            storage[name][1] = DateTime.now()            
            storage[name][2] = abs(storage[name][0]-storage[name][1])
            
            if storage[name][4] == 'count':
            
                storage[name][4] = storage[name][2]
                
        else:
            
            storage[name][0] = DateTime.now()
            storage[name][3] = True
            
    else:
        
        reset(storage, name)        
        
        storage[name] = ['start','stop','current session', True, 'count']
        
        storage[name][0] = DateTime.now()
        
        
    print_counter += 1
    
    if print_counter == print_counter_threshold:
    
        if print_dict(storage, txtfile, xlsxfile) > 0:
        
            print_counter = 0
            
        else:
        
            print_counter -= 1
    
    else:
        
        pass

    time.sleep(2)

