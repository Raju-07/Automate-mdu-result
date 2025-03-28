" Logic for Starting the loop from where it get error "
import random
import os
l = [1,2,3,4,5,6,7,8,9,10]
processed_data = []
file_name = "example.txt"
try:
    if os.path.exists(file_name):
        with open(file_name,"r") as file:
            processed_data = list(map(int,file.read().split()))
except Exception as e:
    print(f"Error is : {e}")
finally:
    with open(file_name,'a') as file:
        for i in l:
                if i in processed_data:
                    continue
                else:      
                    coe = random.uniform(0,1)
                    if coe<0.2:
                        print(i/0)
                    else:
                        print(i)
                        file.write(str(i)+" ")
