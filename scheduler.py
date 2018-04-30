#! /usr/bin/python
import pprint
import pandas as pd

# Create a dictionary for the dummy.xlsx (course -> faculty list)
# For every entry in the dictionary, sort the list based on the seniority
    # need a logic here
# For every entry in the dictionary, assign the faculty member to the slot
    # need a logic here


course_faculty = {}
course_faculty_list = {}
not_assigned = []
preferences = {}
time_slots_dict = {}


# create a map here
def createDictionary():
    xls = pd.ExcelFile('faculty_list.xlsx')
    df1 = pd.read_excel(xls, 'Sheet1')

    faculties = df1['Faculty']
    courses = df1['Course']

    for i in range(len(faculties)):
        if courses[i] not in course_faculty:
            course_faculty[str(courses[i])] = set()
        course_faculty[str(courses[i])].add(" ".join(str(faculties[i]).split()))

def smartSearch(ele, list_collec):
    #list_collec = list_collec.tolist()
    #list_collec = [i.encode('ascii','ignore') for i in list_collec]
    ele = ele.split(" ")
    ele.sort()
    i = 0
    for name in list_collec:
        #name = str(name).encode('utf-8')
        if type(name) == float:
            continue
        x = name.encode('utf-8').split(" ")
        x.sort()
        count = 0
        for a in range(len(ele)):
            for b in range(len(x)):
                if ele[a].lower() == x[b].lower():
                    count = count + 1
        if count >= 2:
            return i
        i = i + 1
    return -1

def seniorityBasedSort():
    # maintain a dict of course to list of faculties
    # SOWK Lead Directory(1)
    # Senior_Lecturers
    # StartForFacultyV5

    # Lead excel sheet
    xls_lead = pd.ExcelFile('Lead.xlsx')
    x_lead = pd.read_excel(xls_lead, 'AY 2017-2018')
    lead_courses = x_lead['Course']
    lead_faculty = x_lead['Lead Instructor VAC']

    # Senior Lecturer Excel Sheet
    xls_lect = pd.ExcelFile('Senior_Lecturers.xlsx')
    x_lect = pd.read_excel(xls_lect, 'Sheet1')
    lect_first = x_lect['First Name']
    lect_last = x_lect['Last Name']
    lect_faculty = []
    for first, last in zip(lect_first, lect_last):
        lect_faculty.append(str(first)+" "+str(last))

    # Faculty based on start day
    xls_fac = pd.ExcelFile('StartForFacultyV5.xlsx')
    x_fac = pd.read_excel(xls_fac, 'Sheet1')

    fac_name = x_fac['Name']
    fac_date = x_fac['Hire Date']

    for key in course_faculty:
        course_faculty_list[key] = []
        temp_sort_list = []
        val = course_faculty[key]

        for facul in val:
            # check in lead
            #if facul in lead_faculty and key == lead_courses[lead_faculty.index(facul)]:
            flag = 0
            ret = smartSearch(facul, lead_faculty)
            if ret>=0 and key == lead_courses[ret]:
                flag=1
                course_faculty_list[key].append(facul)
                continue

            # check in senior
            #if facul in lect_faculty:
            ret = smartSearch(facul, lect_faculty)
            if ret>=0:
                flag=1
                course_faculty_list[key].append(facul)
                continue

            # sort based on hire date
            #if facul in fac_name:
            ret = smartSearch(facul, fac_name)
            if ret>=0:
                flag=1
                temp_sort_list.append(str(fac_date[ret]).split(" ")[1] + "--" + facul)

        if flag == 0:
            not_assigned.append(facul)
        temp_sort_list.sort()

        for ele in temp_sort_list:
            course_faculty_list[key].append(ele.split("--")[1])

    #pprint.pprint(course_faculty_list)
    with open("not_assigned.txt", "w") as myfile:
        for name in not_assigned:
            myfile.write(name + '\n')

def createCombinations(days, timings, flag_any):
    ret_list = []
    if flag_any:
        timings.append('any')
    for day in days:
        for timing in timings:
            for items in time_slots_dict[timing]:
                ret_list.append(str(day)+"---"+str(items))
    return ret_list

def getPreferences():
    fac_pref = pd.ExcelFile('preference.xlsx')
    fac_pref = pd.read_excel(fac_pref, 'Dummy')


    for index, row in fac_pref.iterrows():
        days = []
        timings = []
        flag_any = 0

        if str(row['7:00am-10:00am']) != "nan":
            timings.append('7:00am-10:00am')
        if str(row['10:15am-1:15pm']) != "nan":
            timings.append('10:15am-1:15pm')
        if str(row['4:00pm-7:00pm']) != "nan":
            timings.append('4:00pm-7:00pm')

        if str(row['Any']) != "nan":
            flag_any = 1
            days.extend(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'])
        else:
            if str(row['Mon']) != "nan":
                days.append('Monday')
            if str(row['Tue']) != "nan":
                days.append('Tuesday')
            if str(row['Wed']) != "nan":
                days.append('Wednesday')
            if str(row['Thu']) != "nan":
                days.append('Thursday')
            if str(row['Fri']) != "nan":
                days.append('Friday')
            if str(row['Sat']) != "nan":
                days.append('Saturday')

        name = str(row['Response.1']) + " " + str(row['Response'])

        ret_val = createCombinations(days, timings, flag_any)

        preferences[name] = ret_val

    #print pprint.pprint(preferences)

def setAllTimings():


    '''
    output_excel = pd.ExcelFile("output.xlsx")
    output_excel = pd.read_excel(output_excel, "ALL_ByCourse_8.23")

    item_set = set()

    for item in output_excel['Time (5.19.17)']:
        item_set.add(item)


    for item in item_set:
        if item != "TBA":
            print item
    '''
    time_slots_dict['7:00am-10:00am'] = ['07:00am-09:00am', '08:00am-09:15am', '07:00am-08:15am', '08:45am-10:00am', '08:30am-09:45am', '09:00am-10:15am']
    time_slots_dict['10:15am-1:15pm'] = ['10:15am-11:30am', '10:30am-11:45am', '11:30am-12:45pm', '11:45am-01:00pm', '11:45am-01:45pm', '12:00pm-01:15pm', '10:00am-11:15am', '12:15pm-01:30pm', '09:45am-11:00am']
    time_slots_dict['4:00pm-7:00pm'] = ['04:45pm-06:00pm', '04:15pm-05:30pm', '04:00pm-05:15pm', '04:15pm-06:15pm', '05:00pm-07:00pm', '05:45pm-07:00pm', '03:15pm-04:30pm', '03:45pm-05:00pm','06:30pm-07:45pm']
    time_slots_dict['any'] = ['02:15pm-03:30pm', '06:30pm-08:30pm', '09:15am-11:15am', '02:00pm-03:15pm', '03:00pm-04:15pm', '02:45pm-04:00pm', '02:00pm-04:00pm', '02:30pm-03:45pm', '01:30pm-02:45pm', '01:15pm-02:30pm', '01:00pm-02:15pm']

def prepareSlots(key, output_excel):
    ret_list = []

    for index, row in output_excel.iterrows():
        if row['Course#'] == key:
            timing = row['Time (5.19.17)']
            day = row['Days']
            ret_list.append(str(day) + "---" + str(timing) + "===" + str(index))
    return ret_list



def startAssigning():
    output_excel = pd.ExcelFile("output.xlsx")
    output_excel = pd.read_excel(output_excel, "ALL_ByCourse_8.23")

    for key in course_faculty_list:
        slots = prepareSlots(key, output_excel)
        print key, slots
        #for faculty in course_faculty_list[key]:


if __name__ == "__main__":
    createDictionary()
    seniorityBasedSort()
    setAllTimings()
    getPreferences()
    startAssigning()
