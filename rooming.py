import xlsxwriter


boys_requests = {
    "John": ["Michael", "Joseph", "David"],
    "Michael": ["John", "David", "James"],
    "David": ["John", "Michael", "William"],
    "James": ["Michael", "David", "Joseph"],
    "Robert": ["John", "William", "Joseph"],
    "William": ["Robert", "David", "Joseph"],
    "Joseph": ["Robert", "Thomas", "David"],
    "Thomas": ["Daniel", "Benjamin", "Matthew"],
    "Daniel": ["Thomas", "Benjamin", "James"],
    "Benjamin": ["Charles", "Daniel", "Matthew"],
    "Matthew": ["William", "Daniel", "Benjamin"],
    "Christopher": ["Andrew", "Henry", "Alexander"],
    "Andrew": ["Christopher", "Nicholas", "Alexander"],
    "Nicholas": ["Christopher", "Andrew", "Alexander"],
    "Alexander": ["Christopher", "Andrew", "Nicholas"],
    "William Jr.": ["Charles", "Edward", "Henry"],
    "Charles": ["William Jr.", "Alexander", "Henry"],
    "Edward": ["John", "Charles", "Henry"],
    "Henry": ["William Jr.", "Daniel", "Edward"],
    "Albert": ["David", "James", "Henry"]
}

room_numbers_and_beds = {
    401: 3,
    402: 3,
    403: 3,
    404: 3,
    405: 4,
    406: 4
}

final_rooms = {}




# Main part of the program
if __name__ == "__main__":
    open_rooms = {}
    ##get sorted rooms
    for key, value in room_numbers_and_beds.items():
        if value in open_rooms:
            list1 = list(open_rooms[value])
            list1.append(key)
            open_rooms[value] = list1
        else:
            open_rooms[value] = [key]
   
    

    ##Get Unassigned Boys 
    assigned_boys = {}
    unassigned_boys = []


    count = 1
    for boy1, boy1_requ in boys_requests.items():
        count += 1


        current_room = []
        if boy1 in assigned_boys:
            continue
            
        if(len(open_rooms) > 0):
        
            largest_room_rem = max(open_rooms.keys())
            
            placed_friend = False
            for friend in boy1_requ:
                if len(current_room)+1 == largest_room_rem:
                    break
                if friend in assigned_boys:
                    placed_friend = True
                    continue

                if boy1 in boys_requests.get(friend):
                    current_room.append(friend)
                    
            

            if len(current_room) > 0:
                current_room.append(boy1)

                for n in range(1, largest_room_rem+1):
                    if n in open_rooms.keys() and n >= len(current_room):
                        room_numb = open_rooms[n].pop(0)
                        if(len(open_rooms[n]) == 0):
                            del open_rooms[n]
                        break

                final_rooms[room_numb] = current_room

 
                for b in current_room:
                    assigned_boys[b] = room_numb


            else:
                if(placed_friend):
                    found = False
                    for friend in boy1_requ:
                        if(friend in assigned_boys):
                            room_num = assigned_boys[friend]
                            occupants = final_rooms[room_num]
                            
                            if(len(occupants) < room_numbers_and_beds[room_num]):
                                occupants.append(boy1)
                                final_rooms[room_num] = occupants
                                assigned_boys[boy1] = room_num
                                found = True
                                break

                    if(not found):
                        unassigned_boys.append(boy1)

                else:
                    unassigned_boys.append(boy1)

    
    for boys in boys_requests.keys():
        if(boys not in unassigned_boys and boys not in assigned_boys):
            unassigned_boys.append(boys)

    special_att = []
    for boy1, boy1_requ in boys_requests.items():
        mutual_friend = False
        for friend in boy1_requ:
            if boy1 in boys_requests.get(friend):
                mutual_friend = True
        if(not mutual_friend):
            special_att.append(boy1)
    

    workbook = xlsxwriter.Workbook('Rooms.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "Room Number")
    row = 0
    for room, boys in final_rooms.items():
        col = 0
        row = row+1
        worksheet.write(row, col, room)
        for b in boys:
            col = col+1
            worksheet.write(row, col, b)
        max_cap = room_numbers_and_beds[room]
        while col < max_cap:
            col = col+1
            worksheet.write(row, col, "EMPTY SPOT")
    
    col = 0
    row = row+2
    worksheet.write(row, col, "UNASSIGNED")
    for n in range(1,4):
        worksheet.write(row, n, "REQUESTS")

    

    for unass_b in unassigned_boys:
        col = 0
        row = row+1
        worksheet.write(row, col, unass_b)
        for requ in boys_requests[unass_b]:
            col = col + 1
            worksheet.write(row, col, requ)

    col = 0
    row = row+2
    worksheet.write(row, col, "SPECIAL ATTENTION")
    for n in range(1,4):
        worksheet.write(row, n, "REQUESTS")
    for stud in special_att:
        col = 0
        row = row+1
        worksheet.write(row, col, unass_b)
        for requ in boys_requests[unass_b]:
            col = col + 1
            worksheet.write(row, col, requ)

    workbook.close()
    
