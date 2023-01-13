import gspread
import schedule
import time

# Connecting Python to SGA Data Sheet

sa = gspread.service_account(filename="service_account.json")
sh = sa.open("Sunnyvale Player Stats 2023")

# Setting Each Data and Extrap Sheet to variables for Front Sheet and each player

sga_totals = sh.worksheet("All SGA Total Stats")

john_data = sh.worksheet("John Stats Entry Data")
john_extrap = sh.worksheet("John Stats Extrapolation")

colin_data = sh.worksheet("Colin Stats Entry Data")
colin_extrap = sh.worksheet("Colin Stats Extrapolation")

chris_data = sh.worksheet("Chris Stats Entry Data")
chris_extrap = sh.worksheet("Chris Stats Extrapolation")

cam_data = sh.worksheet("Cam Stats Entry Data")
cam_extrap = sh.worksheet("Cam Stats Extrapolation")

austin_data = sh.worksheet("Austin Stats Entry Data")
austin_extrap = sh.worksheet("Austin Stats Extrapolation")

jeff_data = sh.worksheet("Jeff Player Stat Entry")
jeff_extrap = sh.worksheet("Jeff Stat Extrapolation")

justin_data = sh.worksheet("Justin Stats Entry Data")
justin_extrap = sh.worksheet("Justin Stats Extrapolation")

kevin_data = sh.worksheet("Kevin Stats Entry Data")
kevin_extrap = sh.worksheet("Kevin Stats Extrapolation")

moffa_data = sh.worksheet("Moffa Stats Entry Data")
moffa_extrap = sh.worksheet("Moffa Stats Extrapolation")

ian_data = sh.worksheet("Ian Stat Entry Data")
ian_extrap = sh.worksheet("Ian Stat Extrapolation")

sean_data = sh.worksheet("Sean Stat Entry Data")
sean_extrap = sh.worksheet("Sean Stat Extrapolation")

# Creating Variables
john_last_total_entries = 0
john_event_wins = 0
john_ex_wins = 0

colin_last_total_entries = 0
colin_event_wins = 0
colin_ex_wins = 0

chris_last_total_entries = 0
chris_event_wins = 0
chris_ex_wins = 0

cam_last_total_entries = 0
cam_event_wins = 0
cam_ex_wins = 0

austin_last_total_entries = 0
austin_event_wins = 0
austin_ex_wins = 0

jeff_last_total_entries = 0
jeff_event_wins = 0
jeff_ex_wins = 0

justin_last_total_entries = 0
justin_event_wins = 0
justin_ex_wins = 0

kevin_last_total_entries = 0
kevin_event_wins = 0
kevin_ex_wins = 0

sean_last_total_entries = 0
sean_event_wins = 0
sean_ex_wins = 0

ian_last_total_entries = 0
ian_event_wins = 0
ian_ex_wins = 0

moffa_last_total_entries = 0
moffa_event_wins = 0
moffa_ex_wins = 0


# Defining Functions to update each player
# Function checks to see if the last entries seen are not equal to the current entries number, if different means entry
# sets rounds = to an integer to match player current total entries
# set r = 2, or row = 2 so that it starts at 2 and doesn't count the header when pulling get request
# excount = 0, or exhibition win count = 0, put outside of while loop so that it can update before the data cycles
#













def john_update():
    global evwins, exwins
    john_current_total_entries = int(sga_totals.cell(2, 4).value)
    if john_last_total_entries != john_current_total_entries:
        rounds = int(john_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = john_data.cell(r, 3).value
            win = john_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F2', evwins)
        sga_totals.update('G2', exwins)


# Colin's Function

def colin_update():
    global evwins, exwins
    colin_current_total_entries = int(sga_totals.cell(2, 4).value)
    if colin_last_total_entries != colin_current_total_entries:
        rounds = int(colin_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = colin_data.cell(r, 3).value
            win = colin_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F3', evwins)
        sga_totals.update('G3', exwins)


# Chris' function

def chris_update():
    global evwins, exwins
    chris_current_total_entries = int(sga_totals.cell(2, 4).value)
    if chris_last_total_entries != chris_current_total_entries:
        rounds = int(chris_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = chris_data.cell(r, 3).value
            win = chris_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F4', evwins)
        sga_totals.update('G4', exwins)


# Cam Update Function
def cam_update():
    global evwins, exwins
    cam_current_total_entries = int(sga_totals.cell(2, 4).value)
    if cam_last_total_entries != cam_current_total_entries:
        rounds = int(cam_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2);
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = cam_data.cell(r, 3).value
            win = cam_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F5', evwins)
        sga_totals.update('G5', exwins)


# Austin Update Function
def austin_update():
    global evwins, exwins
    austin_current_total_entries = int(sga_totals.cell(2, 4).value)
    if austin_last_total_entries != austin_current_total_entries:
        rounds = int(austin_current_total_entries)
        r = 2
        excount = 0

        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = austin_data.cell(r, 3).value
            win = austin_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1
                chip_flag


            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F6', evwins)
        sga_totals.update('G6', exwins)


# Jeff Update Function

def jeff_update():
    global evwins, exwins
    jeff_current_total_entries = int(sga_totals.cell(2, 4).value)
    if jeff_last_total_entries != jeff_current_total_entries:
        rounds = int(jeff_current_total_entries)
        r = 2
        excount = 0
        chip_flag = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = jeff_data.cell(r, 3).value
            win = jeff_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F7', evwins)
        sga_totals.update('G7', exwins)


# Justin Update Function

def justin_update():
    global evwins, exwins
    justin_current_total_entries = int(sga_totals.cell(2, 4).value)
    if justin_last_total_entries != justin_current_total_entries:
        rounds = int(justin_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = justin_data.cell(r, 3).value
            win = justin_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F8', evwins)
        sga_totals.update('G8', exwins)


# Kevin Update Function


def kevin_update():
    global evwins, exwins
    kevin_current_total_entries = int(sga_totals.cell(2, 4).value)
    if kevin_last_total_entries != kevin_current_total_entries:
        rounds = int(kevin_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = kevin_data.cell(r, 3).value
            win = kevin_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F9', evwins)
        sga_totals.update('G9', exwins)


# Moffa Update Function

def moffa_update():
    global evwins, exwins
    moffa_current_total_entries = int(sga_totals.cell(2, 4).value)
    if moffa_last_total_entries != moffa_current_total_entries:
        rounds = int(moffa_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = kevin_data.cell(r, 3).value
            win = kevin_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F10', evwins)
        sga_totals.update('G10', exwins)


# Ian Update Function

def ian_update():
    global evwins, exwins
    ian_current_total_entries = int(sga_totals.cell(2, 4).value)
    if ian_last_total_entries != ian_current_total_entries:
        rounds = int(ian_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = ian_data.cell(r, 3).value
            win = ian_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F11', evwins)
        sga_totals.update('G11', exwins)


# Sean Update Function
def sean_update():
    global evwins, exwins
    sean_current_total_entries = int(sga_totals.cell(2, 4).value)
    if sean_last_total_entries != sean_current_total_entries:
        rounds = int(sean_current_total_entries)
        r = 2
        excount = 0
        while rounds > 0:
            time.sleep(2)
            exwins = 0
            evcount = 0
            evwins = 0
            exflag = 0
            wflag = 0
            am_flag = 0
            ai_flag = 0
            mm_flag = 0
            mi_flag = 0
            junm_flag = 0
            juni_flag = 0
            julm_flag = 0
            juli_flag = 0
            augm_flag = 0
            augi_flag = 0
            chip_flag = 0

            match_type = sean_data.cell(r, 3).value
            win = sean_data.cell(r, 9).value

            if ((match_type == ("Exhibition")) and (win == "TRUE")):
                excount = excount + 1
            elif match_type == ("April Major Event"):
                am_flag = am_flag + 1
            elif match_type == ("April Invitational Event"):
                ai_flag = ai_flag + 1
            elif match_type == ("May Major Event"):
                mm_flag = mm_flag + 1
            elif match_type == ("May Invitational Event"):
                mi_flag = mi_flag + 1
            elif match_type == ("June Major Event"):
                junm_flag = junm_flag + 1
            elif match_type == ("June Invitational Event"):
                juni_flag = juni_flag + 1
            elif match_type == ("July Major Event"):
                julm_flag = julm_flag + 1
            elif match_type == ("July Invitational Event"):
                juli_flag = juli_flag + 1
            elif match_type == ("August Major Event"):
                augm_flag = augm_flag + 1
            elif match_type == ("August Invitational Event"):
                augi_flag = augi_flag + 1
            elif match_type == ("Championship"):
                chip_flag = chip_flag + 1

            if win == ("TRUE"):
                wflag = wflag + 1

            if ((wflag == 1) and (exflag == 1)):
                excount = excount + 1

            if ((wflag == 1) and (am_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (ai_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (mi_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (junm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juni_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (julm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (juli_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augm_flag == 1)):
                evcount = evcount + 1
            elif ((wflag == 1) and (augi_flag == 1)):
                evcount = evcount + 1

            r = r + 1
            exwins = excount
            evwins = evcount
            rounds = rounds - 1

        sga_totals.update('F12', evwins)
        sga_totals.update('G12', exwins)
        print(evwins)
        print(exwins)


# Schedule Job that runs all functions, and have it run every 2 minutes


def steady():
    john_update()
    time.sleep(120)
    colin_update()
    time.sleep(120)
    cam_update()
    time.sleep(120)
    chris_update()
    time.sleep(120)
    austin_update()
    time.sleep(120)
    jeff_update()
    time.sleep(120)
    justin_update()
    time.sleep(120)
    moffa_update()
    time.sleep(120)
    kevin_update()
    time.sleep(120)
    ian_update()
    time.sleep(120)
    sean_update()


steady()








