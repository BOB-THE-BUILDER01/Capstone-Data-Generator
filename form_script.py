import xlsxwriter as ex
import random as rd
import os

# Open / Create the Excel Sheet
cwd = os.path.dirname(os.path.realpath(__file__))
book = ex.Workbook(os.path.join(os.path.dirname(os.path.abspath(__file__)),"bigbrain.xlsx"))
sheet = book.add_worksheet()

# Define the Colmun Names
sheet.write(0, 0, "Serial No.")
sheet.write(0, 1, "Description of the situation the animal is in")
sheet.write(0, 2, "Animal type")
sheet.write(0, 3, "Type of Situation / Category")

# Define the Categories to write about.
row = 1
animal_type = ["Bird", "Cat", "Dog", "Rabbit", "Squirrel", "Cattle", "Pig", "Pigeon", "Hen", "Donkey", "Elephant", "Horse", "Carnivorous", "Monkey", "Peacock", "Snake"]
situation = ["Adoption", "Neutering", "Cruelty_Complaint", "Illegal_Breeding", "Illegal_Pet_Shop", "Illegal_Possession_of_Exotic_Animal", "Injured_Sick", "Rescue", "Lost_Displaced_Abandoned", "Weak_Underfed_Malnourished"]

# All the statements from the different categories.
normal_opening = open(cwd + "\\starters.txt", 'r')
illegal_opening = open(cwd + "\\illegal_starters.txt", 'r')
location =  open(cwd + "\\locations.txt", 'r')
extra_info =  open(cwd + "\\extras.txt", 'r')

pet = False
wild = False

selected_opening_normal = normal_opening.readlines()
selected_opening_illegal = illegal_opening.readlines()
selected_location = location.readlines()
selected_extra = extra_info.readlines()

animal_dict = dict()
situation_dict = dict()

for i in range(0, 20000):
    complaint = ""
    # Randomly Select all the different Variables
    while(True):
        at = rd.randrange(0, len(animal_type), 1)
        if(animal_type[at] in animal_dict):
            if(animal_dict[animal_type[at]] > 2000):
                continue
            else:
                animal_dict[animal_type[at]] += 1
                break
        
        else:
            animal_dict[animal_type[at]] = 1
            break
    
    extra = rd.randrange(0,3)
    pet = True if at <= 4 else False
    an_info = None

    if(pet):
        # Select Pet Situation
        sit = -1
        while(True):
            sit = rd.randrange(0, len(situation), 1)
            if(sit == 5):
                continue
            elif(situation[sit] in situation_dict):
                if(situation_dict[situation[sit]] > 2000):
                    continue
                else:
                    situation_dict[situation[sit]] += 1
                    break
            else:
                situation_dict[situation[sit]] = 1
                break
    else:
        # Select Wild / Cattle / Pest Situation
        sit = rd.randrange(2, len(situation), 1)
        an_info = open(cwd + "\\extra\\" + animal_type[at] + "-info.txt", 'r')

    # FORMULATE THE COMPLAINT -> Categories defined here https://docs.google.com/document/d/1_5zIp5IzvX_V5LfekVkBloVqMf6O72lifNHryzizIPk/edit?usp=sharing
    sitfile = open(cwd + "\\situations\\" + situation[sit] + ".txt", 'r')

    # Adding the Opening and the Location - Change the opening location based on type of situation for realism.
    if(sit >= 3 and sit < 6):
        complaint += selected_opening_illegal[rd.randrange(0, len(selected_opening_illegal), 1)].split('\n')[0]
    else:
        complaint += selected_opening_normal[rd.randrange(0, len(selected_opening_normal), 1)].split('\n')[0]
    
    complaint += selected_location[rd.randrange(0, len(selected_location), 1)].split('\n')[0]

    # Adding the situation in which the animal is in
    selected_sit = sitfile.readlines()
    complaint += selected_sit[rd.randrange(0, len(selected_sit), 1)].split("\n")[0]

    # Add more specific information about Wild / Cattle animals towards the end if True
    if(an_info != None):
        selected_info = an_info.readlines()
        complaint += selected_info[rd.randrange(0, len(selected_info))].split('\n')[0]
    
    # Adding extra peppered information in the beginning or the end if True
    if(extra > 0):
        if(extra == 1):
            complaint = selected_extra[rd.randrange(0, len(selected_extra), 1)].split('\n')[0] + complaint
        else:
            complaint += selected_extra[rd.randrange(0, len(selected_extra), 1)].split('\n')[0]

    # Change the $'s to animal names
    complaint = complaint.replace("$", animal_type[at])
    
    # Add the row to the sheet
    sheet.write(row, 0, str(i))
    sheet.write(row, 1, complaint)
    sheet.write(row, 2, animal_type[at])
    sheet.write(row, 3, situation[sit])
    row += 1

    sitfile.close()
    if(an_info != None):
        an_info.close()

# Close once Done -> Script won't work if you do not do this step
book.close()

location.close()
extra_info.close()
normal_opening.close()
illegal_opening.close()
