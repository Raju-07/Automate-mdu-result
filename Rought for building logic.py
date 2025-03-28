" Logic for Starting the loop from where it get error "

import random
import os
l = [1,2,3,4,5,6,7,8,9,10]
list2 = []
file_name = "example.txt"
try:
    if os.path.exists(file_name):
        with open(file_name,"r") as file:
            pre_data = file.read()
            list2 = pre_data.split()
            list2 = list(map(int,list2))
    else:
        with open(file_name,"w") as file:
            pass
except Exception as e:
    print(f"Error is : {e}")
finally:
    for i in l:
            if i in list2:
                pass
            else:      
                with open(file_name,"a") as file2:
                    coe = random.uniform(0,1)
                    if coe<0.2:
                        print(i/0)
                    else:
                        print(i)
                        file2.write(str(i)+" ")



# unique_data = [x for x in l if x not in list2] + [int(x) for x in list2 if x not in l]
# final_data = list(set(unique_data))





