import pandas

# Read dataframe values, convert to list for database
ccc_df = pandas.read_excel(io = "Database.xlsx", sheet_name = "Records")
ccc_df_list = ccc_df.values.tolist()

# Keep a dictionary of the column titles/indices per entry
headers = list(ccc_df.columns)
ccc_index_dict = {headers[i]: i for i in range(len(headers))}

# Lists for the performer objects, the output list, and declined list
ccc_performers = []
ccc_outputs = [['Name: ', 'Teacher: ', 'Month: ', 'Timeslot: ', 'Attendance: ', 'Eligibility: ']]
ccc_declined = [['Name: ', 'Teacher: ', 'Requested Month: ', 'Requested Timeslot: ', 'Eligibility: ', 'Attendance: ', 'Comment: ']]

# Read in excepted schools to the "1 per school" limit
ccc_school_excepts_df = pandas.read_excel(io = "Database.xlsx", sheet_name = "School Exceptions")
ccc_school_excepts = ccc_school_excepts_df.values.tolist()

# Read in the months for this year
planned_month_input = pandas.read_excel(io = "Database.xlsx", sheet_name = "Planned Months").values.tolist()

class MonthTime:
    """To keep track of timings and the amount of available slots each month."""
    TIME_LIMIT, months = 155, {}

    # Exception: 20 minute slots, anything >= 20 minutes can pretty much do RTP in the next slot

    # 15 : 4
    max_times = {5: 3, 10: 4, 15: 5, 20: 3, 25: 2, 30: 1}

    def __init__(self, name):
        """Keep track of attributes, store month in list"""
        self.name = name
        self.time, self.times_available = MonthTime.TIME_LIMIT, dict(MonthTime.max_times)
        MonthTime.months[name] = self

    def available(month, req):
        """Check availability."""
        m = MonthTime.months[month]

        # Does the time fit in, and are there slots for that time available in the month
        return m.time >= req and m.times_available[req] > 0

    def mod(month, req):
        """Modify month given a slot being used."""
        m = MonthTime.months[month]

        # Decrease total time and slots available
        m.time -= req
        m.times_available[req] -= 1

class Performer:
    """A class for representing a performer and timeslot."""

    # Store timeslots as tuples - performers in one "range" of slots
    # are promoted to the next (ex: you could go 20 to 25 or 20 to 30)
    times = [range(5,10), range(10,15), range(15,20), range(20, 30), range(30,40)]

    # Who can perform, then what months available
    ineligibility_code = "NS"
    planned_months = [month[0] for month in planned_month_input]

    # Limits and exceptions
    SCHOOL_LIMIT, TIME_LIMIT = 1, 155
    school_dict = dict(ccc_school_excepts)
    
    # Tracking constraints
    schools_per_month = {m: {} for m in planned_months}
    for m in planned_months:
        MonthTime(m)
    
    flexibility_queue = []
    
    def __init__(self, name, school, eligibility, month1, month2, req_time, attendance):
        """Initialize performer with the appropriate attributes.
        Includes keeping track of eligibility, timeslot, comments, etc. 
        """
        # Set attributes
        self.name, self.school, self.attendance = name, school, attendance
        self.month1, self.month2 = month1, month2
        self.req_time = req_time
        self.eligibility = eligibility
        self.comments = []
 
    def assign(self):
        """Assigns slot if validated, else declines."""

        msl, msa = self.month_school_limit, self.month_slot_available
        self.declined = False

        # All reasons that this doesn't work, in order
        for month in self.month1, self.month2:
            if self.eligibility == Performer.ineligibility_code:
                self.comments.append("The performer's ability to perform in this half is uncertain")
                self.declined = True
                return
            elif month == "None":
                Performer.flexibility_queue.append(self)
                return
            elif month not in Performer.planned_months:
                self.comments.append(month + " is not a valid request")
            elif not msl(month):
                self.comments.append("For " + month + ", the school limit for " + self.school + " is reached")
            elif not msa(month):
                self.comments.append("For " + month + ", the " + str(self.req_time) + "-minute slot is unavailable")
            else:
                # Success!
                self.set_month_and_time(month)
                return
        
        self.declined = True

    def month_slot_available(self, month):
        """Checks whether month has the appropriate slot."""
        return MonthTime.available(month, self.req_time)
    
    def month_school_limit(self, month):
        """Checks whether the number of students in one school
        for a given month does not exceed limit.
        Note: Limit can be the standard, or non-standard if the school is in exceptions dictionary."""
        return Performer.schools_per_month[month].get(self.school, 0) < Performer.school_dict.get(self.school, Performer.SCHOOL_LIMIT)
    
    def set_month_and_time(self, month):
        """Modify the appropriate records for school and timeslot in the chosen month."""

        Performer.schools_per_month[month][self.school] = Performer.schools_per_month[month].get(self.school, 0) + 1

        MonthTime.mod(month, self.req_time)

        # Record final results
        self.this_month, self.this_time = month, self.req_time
    
    def special_assign(self):
        """Specially assigns performers who have no preference."""

        # Iterate through each month and find something that works
        for month in Performer.planned_months:
            if self.month_school_limit(month) and self.month_slot_available(month):
                self.declined = False
                self.set_month_and_time(month)
                return
        
        self.declined = True
        self.comments.append("Performer had no preference and no months were logistically available")


# Create performers from list of data and do assignment
for p in ccc_df_list:
    if not pandas.isnull(p[0]):
        pname, pschool = p[ccc_index_dict['Name: ']], p[ccc_index_dict['Teacher: ']]
        pmonth1, pmonth2 = p[ccc_index_dict["First choice: "]], p[ccc_index_dict['Second choice: ']]
        preq_time = p[ccc_index_dict['Requested time: ']]
        pattendance = p[ccc_index_dict['Attendance: ']]
        peligibility = p[ccc_index_dict["Eligibility"]]

        pf = Performer(name = pname, school = pschool, month1 = pmonth1, month2 = pmonth2, req_time = preq_time, attendance = pattendance, eligibility = peligibility)
        ccc_performers.append(pf)

# Sort from most to least attendance
ccc_performers.sort(key = lambda p: p.attendance, reverse = True)

# Set up month-based output dictionary and counts for declined/assigned performers
ccc_month_output = {}

for month in Performer.planned_months:
    ccc_month_output[month] = []

count_assigned, count_declined = 0, 0

# Do assignment
for p in ccc_performers:
    p.assign()

# Do special assignment
for p in Performer.flexibility_queue:
    p.special_assign()

# Record results
for p in ccc_performers:
    # Sort into declined and accepted, record into month tables
    if p.declined:
        for m, c in zip((p.month1, p.month2), p.comments * 2):
            ccc_declined.append([p.name, p.school, m, p.req_time, p.eligibility, p.attendance, c])
        count_declined += 1
    else:
        output_lst = [p.name, p.school, p.this_month, p.this_time, p.attendance, p.eligibility]
        ccc_outputs.append(output_lst)
        ccc_month_output[p.this_month].append(output_lst[0:2] + output_lst[3:])
        count_assigned += 1

# Add blank row and final counts to output/declined sheets
ccc_outputs.append([""] * len(ccc_outputs[0]))
ccc_outputs.append(['Count: ' + str(count_assigned)] + [""] * (len(ccc_outputs[0]) - 1))

ccc_declined.append([""] * len(ccc_declined[0]))
ccc_declined.append(['Count: ' + str(count_declined)] + [""] * (len(ccc_declined[0]) - 1))

# Write to data frames
ccc_outputs_df = pandas.DataFrame(ccc_outputs[1:], columns = ccc_outputs[0])
ccc_declined_df = pandas.DataFrame(ccc_declined[1:], columns = ccc_declined[0])

# Set up Excel sheet and format
xlwriter = pandas.ExcelWriter('Output.xlsx', engine = 'xlsxwriter')
format = xlwriter.book.add_format({'text_wrap': True})

# Write to sheets and finish
ccc_outputs_df.to_excel(xlwriter, sheet_name = 'Outputs')
ccc_declined_df.to_excel(xlwriter, sheet_name = 'Declined')

# Columns for each month sheet
ccc_month_cols = ccc_outputs[0][0:2] + ccc_outputs[0][3:]

# Write to each month sheet and finish
for month in ccc_month_output:
    ccc_month_df = pandas.DataFrame(ccc_month_output[month], columns = ccc_month_cols)
    ccc_month_df.to_excel(xlwriter, sheet_name = month + " 2021")

xlwriter.save()