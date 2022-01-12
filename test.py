import pandas as pd
pd.options.mode.chained_assignment = None

regfile = 'paniitmainreg.xlsx'
menfile = 'paniitmainmen.xlsx'
womenfile = 'paniitmainwomen.xlsx'
teamfile = 'paniitmainteam.xlsx'
day_count = 5
leaderboard_needed = 't'

reg = pd.read_excel(regfile, sheet_name = 'reg')
reg_men = reg['Strava Username'][reg['Your Category'] == 'Men: Join https://www.strava.com/clubs/1009649'].tolist()
reg_women = reg['Strava Username'][reg['Your Category'] == 'Women: Join https://www.strava.com/clubs/1009650'].tolist()

dm = pd.read_excel(menfile, sheet_name = 'distance')
sm = pd.read_excel(menfile, sheet_name = 'speed')
em = pd.read_excel(menfile, sheet_name = 'elevation')
dw = pd.read_excel(womenfile, sheet_name = 'distance')
sw = pd.read_excel(womenfile, sheet_name = 'speed')
ew = pd.read_excel(womenfile, sheet_name = 'elevation')

iits = pd.read_excel(teamfile, sheet_name = 'iits')['IIT'].tolist()

leaderboards = [dm, sm, em, dw, sw, ew]
replacement = [['Distance', ' km'], ['Avg. Speed', ' km/h'], ['Elev. Gain', ' m']]
count = 1
for board in leaderboards:
    board['Registered'] = None
    for i in board.index:
        for r in replacement:
            if ',' in board[r[0]][i]: board[r[0]][i] = board[r[0]][i].replace(',', '')
            board[r[0]][i] = float(board[r[0]][i].replace(r[1], ''))
        if count<=3: reg_list = reg_men
        else: reg_list = reg_women
        if board['Athlete'][i] in reg_list: 
            iit = reg['Name of your Institute'][reg['Strava Username'] == board['Athlete'][i]].sort_values(ignore_index = True)
            if iit[0] != 'Indian Institute of Technology (IIT), Kanpur': board['Registered'][i] = 'Yes'
        else: board['Registered'][i] = 'No'
    count += 1    

dm = dm.loc[dm['Registered'] == 'Yes'].sort_values(by = ['Distance'], ascending = False, ignore_index = True)
sm = sm.loc[sm['Registered'] == 'Yes'].sort_values(by = ['Avg. Speed'], ascending = False, ignore_index = True)
em = em.loc[em['Registered'] == 'Yes'].sort_values(by = ['Elev. Gain'], ascending = False, ignore_index = True)
dw = dw.loc[dw['Registered'] == 'Yes'].sort_values(by = ['Distance'], ascending = False, ignore_index = True)
sw = sw.loc[sw['Registered'] == 'Yes'].sort_values(by = ['Avg. Speed'], ascending = False, ignore_index = True)
ew = ew.loc[ew['Registered'] == 'Yes'].sort_values(by = ['Elev. Gain'], ascending = False, ignore_index = True)
sm = sm.loc[sm['Distance'] >= 100*(day_count/7)].sort_values(by = ['Avg. Speed'], ascending = False, ignore_index = True)
sw = sw.loc[sw['Distance'] >= 20*(day_count/7)].sort_values(by = ['Avg. Speed'], ascending = False, ignore_index = True)

men_dl = dm['Athlete'].tolist()
men_sl = sm['Athlete'].tolist()
men_el = em['Athlete'].tolist()
men_l = men_dl + list(set(men_sl) - set(men_dl))
men_l = men_l + list(set(men_el) - set(men_l))
omp = []
for man in men_l:
    points = 0
    if man in men_dl: points += men_dl.index(man)
    else: points += 100
    if man in men_sl: points += men_sl.index(man)
    else: points += 100
    if man in men_el: points += men_el.index(man)
    else: points += 100
    points = 300 - points
    omp.append(points)    
om = pd.DataFrame({'Athlete': men_l, 'Points': omp})
om = om.sort_values(by = ['Points'], ascending = False, ignore_index = True)

women_dl = dw['Athlete'].tolist()
women_sl = sw['Athlete'].tolist()
women_el = ew['Athlete'].tolist()
women_l = women_dl + list(set(women_sl) - set(women_dl))
women_l = women_l + list(set(women_el) - set(women_l))
owp = []
for woman in women_l:
    points = 0
    if woman in women_dl: points += women_dl.index(woman)
    else: points += 100
    if woman in women_sl: points += women_sl.index(woman)
    else: points += 100
    if woman in women_el: points += women_el.index(woman)
    else: points += 100
    points = 300 - points
    owp.append(points)    
ow = pd.DataFrame({'Athlete': women_l, 'Points': owp})
ow = ow.sort_values(by = ['Points'], ascending = False, ignore_index = True)

mens_leaderboard = pd.DataFrame({})
mens_leaderboard['Athlete (Overall)'] = om['Athlete']
mens_leaderboard['Points'] = om['Points']
mens_leaderboard['Athlete (Distance)'] = dm['Athlete']
mens_leaderboard['Distance'] = dm['Distance']
mens_leaderboard['Athlete (Avg. Speed)'] = sm['Athlete']
mens_leaderboard['Avg. Speed'] = sm['Avg. Speed']
mens_leaderboard['Athlete (Elev. Gain)'] = em['Athlete']
mens_leaderboard['Elev. Gain'] = em['Elev. Gain']

womens_leaderboard = pd.DataFrame({})
womens_leaderboard['Athlete (Overall)'] = ow['Athlete']
womens_leaderboard['Points'] = ow['Points']
womens_leaderboard['Athlete (Distance)'] = dw['Athlete']
womens_leaderboard['Distance'] = dw['Distance']
womens_leaderboard['Athlete (Avg. Speed)'] = sw['Athlete']
womens_leaderboard['Avg. Speed'] = sw['Avg. Speed']
womens_leaderboard['Athlete (Elev. Gain)'] = ew['Athlete']
womens_leaderboard['Elev. Gain'] = ew['Elev. Gain']

tdl = men_dl + list(set(men_el) - set(men_dl))
tdl = tdl + list(set(women_dl) - set(tdl))
tdl = tdl + list(set(women_el) - set(tdl))
tsl = tdl + list(set(men_sl) - set(tdl))
tsl = tsl + list(set(women_sl) - set(tsl))
tel = tdl

tdd = []
for leader in tdl:
    if leader in men_dl: distance_covered = dm['Distance'].tolist()[men_dl.index(leader)]
    elif leader in men_el: distance_covered = em['Distance'].tolist()[men_el.index(leader)]
    elif leader in women_dl: distance_covered = dw['Distance'].tolist()[women_dl.index(leader)]
    else: distance_covered = ew['Distance'].tolist()[women_el.index(leader)]
    tdd.append(distance_covered)
team_distance = pd.DataFrame({'Athlete': tdl, 'Distance': tdd})
team_distance = team_distance.sort_values(by = ['Distance'], ascending = False, ignore_index = True)

tss = []
tsd = []
for leader in tsl:
    if leader in men_dl: 
        speed_reached = dm['Avg. Speed'].tolist()[men_dl.index(leader)]
        distance_covered = dm['Distance'].tolist()[men_dl.index(leader)]
    elif leader in men_el: 
        speed_reached = em['Avg. Speed'].tolist()[men_el.index(leader)]
        distance_covered = em['Distance'].tolist()[men_el.index(leader)]
    elif leader in women_dl: 
        speed_reached = dw['Avg. Speed'].tolist()[women_dl.index(leader)]
        distance_covered = dw['Distance'].tolist()[women_dl.index(leader)]
    elif leader in women_el: 
        speed_reached = ew['Avg. Speed'].tolist()[women_el.index(leader)]
        distance_covered = ew['Distance'].tolist()[women_el.index(leader)]
    elif leader in men_sl: 
        speed_reached = sm['Avg. Speed'].tolist()[men_sl.index(leader)]
        distance_covered = sm['Distance'].tolist()[men_sl.index(leader)]
    else:
        speed_reached = sw['Avg. Speed'].tolist()[women_sl.index(leader)]
        distance_covered = sw['Distance'].tolist()[women_sl.index(leader)]
    tss.append(speed_reached)
    tsd.append(distance_covered)
team_speed = pd.DataFrame({'Athlete': tsl, 'Avg. Speed': tss, 'Distance': tsd})
team_speed = team_speed.sort_values(by = ['Avg. Speed'], ascending = False, ignore_index = True)

tee = []
for leader in tel:
    if leader in men_dl: elevation_gained = dm['Elev. Gain'].tolist()[men_dl.index(leader)]
    elif leader in men_el: elevation_gained = em['Elev. Gain'].tolist()[men_el.index(leader)]
    elif leader in women_dl: elevation_gained = dw['Elev. Gain'].tolist()[women_dl.index(leader)]
    else: elevation_gained = ew['Elev. Gain'].tolist()[women_el.index(leader)]
    tee.append(elevation_gained)
team_elevation = pd.DataFrame({'Athlete': tel, 'Elev. Gain': tee})
team_elevation = team_elevation.sort_values(by = ['Elev. Gain'], ascending = False, ignore_index = True)

td = pd.DataFrame({'IIT': [], 'Distance': []})
ts = pd.DataFrame({'IIT': [], 'Avg. Speed': [], '600 km': [], 'Categorise': []})
te = pd.DataFrame({'IIT': [], 'Elev. Gain': []})
tp = pd.DataFrame({'IIT': [], 'Points': []})
td['IIT'] = iits
ts['IIT'] = iits
te['IIT'] = iits
tp['IIT'] = iits

for i in td.index:
    distance = 0
    for athlete in team_distance['Athlete'].tolist():
        iit = reg['Name of your Institute'][reg['Strava Username'] == athlete].sort_values(ignore_index = True)
        if td['IIT'][i] == iit[0]:
            distance += team_distance['Distance'].tolist()[team_distance['Athlete'].tolist().index(athlete)]
    td['Distance'][i] = distance
td = td.sort_values(by = ['Distance'], ascending = False, ignore_index = True)

for i in ts.index:
    distance_remaining = 600*(day_count/7)
    time = 0
    for athlete in team_speed['Athlete'].tolist():
        iit = reg['Name of your Institute'][reg['Strava Username'] == athlete].sort_values(ignore_index = True)
        if ts['IIT'][i] == iit[0]:
            d = team_speed['Distance'].tolist()[team_speed['Athlete'].tolist().index(athlete)]
            s = team_speed['Avg. Speed'].tolist()[team_speed['Athlete'].tolist().index(athlete)]
            if distance_remaining <= d:
                time = time + distance_remaining/s
                distance_remaining = 0
            else:
                time = time + d/s
                distance_remaining = distance_remaining - d
    if distance_remaining == 0:
        ts['Avg. Speed'][i] = 600*(day_count/7)/time
        ts['600 km'][i] = 'Yes'
        ts['Categorise'][i] = ts['Avg. Speed'][i] + 1000
    else:
        distance_remaining = 300*(day_count/7)
        time = 0
        for athlete in team_speed['Athlete'].tolist():
            iit = reg['Name of your Institute'][reg['Strava Username'] == athlete].sort_values(ignore_index = True)
            if ts['IIT'][i] == iit[0]:
                d = team_speed['Distance'].tolist()[team_speed['Athlete'].tolist().index(athlete)]
                s = team_speed['Avg. Speed'].tolist()[team_speed['Athlete'].tolist().index(athlete)]
                if distance_remaining <= d:
                    time = time + distance_remaining/s
                    distance_remaining = 0
                else:
                    time = time + d/s
                    distance_remaining = distance_remaining - d
        if distance_remaining == 0:
            ts['Avg. Speed'][i] = 300*(day_count/7)/time
            ts['600 km'][i] = 'No'
            ts['Categorise'][i] = ts['Avg. Speed'][i]
        else:
            ts['Avg. Speed'][i] = 0
            ts['600 km'][i] = 'No'
            ts['Categorise'][i] = 0
ts = ts.sort_values(by = ['Categorise'], ascending = False, ignore_index = True)    

for i in te.index:
    elevation = 0
    for athlete in team_elevation['Athlete'].tolist():
        iit = reg['Name of your Institute'][reg['Strava Username'] == athlete].sort_values(ignore_index = True)
        if te['IIT'][i] == iit[0]:
            elevation += team_elevation['Elev. Gain'].tolist()[team_elevation['Athlete'].tolist().index(athlete)]
    te['Elev. Gain'][i] = elevation
te = te.sort_values(by = ['Elev. Gain'], ascending = False, ignore_index = True)

for i in tp.index:
    team_points_list = [25, 18, 15, 12, 10, 8, 6, 4, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    points = 0
    d_index = td['IIT'].tolist().index(tp['IIT'][i])
    if td['Distance'].tolist()[d_index] > 0: points += team_points_list[d_index]
    s_index = ts['IIT'].tolist().index(tp['IIT'][i])
    if ts['Avg. Speed'].tolist()[s_index] > 0: points += team_points_list[s_index]
    e_index = te['IIT'].tolist().index(tp['IIT'][i])
    if te['Elev. Gain'].tolist()[e_index] > 0: points += team_points_list[e_index]
    tp['Points'][i] = points
tp = tp.sort_values(by = ['Points'], ascending = False, ignore_index = True)

team_leaderboard = pd.DataFrame({})
team_leaderboard['IIT (Overall)'] = tp['IIT']
team_leaderboard['Points'] = tp['Points']
team_leaderboard['IIT (Distance)'] = td['IIT']
team_leaderboard['Distance'] = td['Distance']
team_leaderboard['IIT (Avg. Speed)'] = ts['IIT']
team_leaderboard['Avg. Speed'] = ts['Avg. Speed']
team_leaderboard['600 km'] = ts['600 km']
team_leaderboard['IIT (Elev. Gain)'] = te['IIT']
team_leaderboard['Elev. Gain'] = te['Elev. Gain']

if leaderboard_needed == 'm':
    mens_leaderboard.index += 1
    mens_leaderboard.to_clipboard()
if leaderboard_needed == 'w':
    womens_leaderboard.index += 1
    womens_leaderboard.to_clipboard()
if leaderboard_needed == 't':
    team_leaderboard.index += 1
    team_leaderboard.to_clipboard()