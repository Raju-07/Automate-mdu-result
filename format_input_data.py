registration_no = """23120516      
 """.split() #Enter All registraion Number with space 

roll_no = """607
 """.split() #Enter All Roll no with space

# #convert to int and zip

data = list(zip(map(int,registration_no),map(int,roll_no)))
print(data)
