import os
import random
import random
current_row = 4
def tracker_row():
    if os.path.exists("demo.txt"):
        with open("demo.txt","r") as file:
            row = file.readline().strip()  
    else:
        with open("demo.txt","w") as file:
            file.write(str(current_row))
        row = str(current_row)
    return row

data = tracker_row()
current_row = int(tracker_row())
for _ in range(10):  # Example: 10 iterations
    if random.random() > 0.2:  # 80% chance of success
        current_row += 1
        with open("demo.txt", "w") as file:
            file.write(str(current_row))
        print(f"Success: current_row updated to {current_row}")
    else:
        print("Error occurred, current_row not updated.")