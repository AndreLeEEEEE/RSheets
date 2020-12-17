import openpyxl

def parts(requests):
    loc_name = input("Enter the location name")
    pass

def main():
    print("Welcome Rick Sr.")

    request_dict = {}  # Key is Location, Value is list pair of Part No and Qty
    loop_cond = False
    while(not loop_cond):
        locs = input("Enter in the amount of locations you want parts sent to: ")
        if not locs.isdigit():
            print("Error: Your input wasn't a digit\n")
        elif int(locs) is 0:
            print("You have indicated that you wish to send parts to zero locations.\nEnding program.")
            break
        else:
            locs = int(locs)
            loop_cond = True

    if loop_cond is False:
        pass
    else:
        for index in range(locs):
            index += 1
            print("For location % 2d" %(index))
            parts(request_dict)

main()

#wb_obj = Workbook()  # Create a new excel workbook
#sheet_obj = wb_obj.active  # Get the current (only) sheet and store it into an object

#wb_obj.save([Name])