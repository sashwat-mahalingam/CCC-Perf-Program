import pandas

def attendance(filename):
    """A function that counts up the attendance from the given filename."""

    # Read in attendance from the given file, start up the tracker
    attendance_df = pandas.read_excel(io=filename, sheet_name="Attendance")
    attendance_list = attendance_df.values.tolist()
    attendance_list = attendance_list[1:]
    counts = []

    # For each entry, note the first and last time and sum the non-blank attendance slots
    for entry in attendance_list:
        if not pandas.isnull(entry[0]):
            counts.append([str(entry[2]) + " " + str(entry[1]),
                           sum([int(not pandas.isnull(x)) for x in entry[3:15]])])
    
    # Write to output
    counts_df = pandas.DataFrame(
        counts, columns=['Name: ', 'Attendance Count: '])
    xlwriter = pandas.ExcelWriter(
        'Attendance Counts.xlsx', engine='xlsxwriter')
    counts_df.to_excel(xlwriter, sheet_name='Counts')
    worksheet1 = xlwriter.sheets['Counts']
    worksheet1.set_column(1, 2, 30)
    xlwriter.save()

def flatten_dataset():
    """Compile all data into one sheet."""

    # Mapping all the months to numbers
    months = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']
    months = {i + 1: months[i] for i in range(12)}

    # Reading in the attendance compiled earlier
    ccc_attendance_list = pandas.read_excel(
        io='Attendance Counts.xlsx', sheet_name='Counts').values.tolist()
    
    # Cut out the number indices from the attendance list
    ccc_attendance_list = [x[1:] for x in ccc_attendance_list]

    # Remove all the numbers from people's names
    for p in ccc_attendance_list:
        p[0] = "".join([p[0][i] for i in range(len(p[0]))
                        if p[0][i] not in [str(x) for x in range(10)]]).strip()

    # Read in last year's performances
    last_year = pandas.read_excel(
        io="Scheduling.xlsx", sheet_name='2020').values.tolist()
    
    # Extracting the names (with surrounding whitespace removed), months, and timeslots from last year
    last_year = [[(x[1]).strip(), months[x[0].month], (x[3])] for x in last_year]

    # Extracting names (without surrounding whitespace), school, requested months, requested timeslots
    requests = [[x[0].strip()] + x[1:5] for x in pandas.read_excel(
        io="Requests.xlsx", sheet_name="Responses").values.tolist()[1:] if not pandas.isnull(x[0])]

    # Figure out the times requested, the months (if any were requested)
    for x in requests:
        x[-1] = max([t for t in [5, 10, 15, 20, 25, 30]
                 if (str(t) + " mts" in x[-1])])

        x[2], x[3] = [m for m in ['January', 'February', 'March', 'April', 'May', 'June'] if (
            m in x[2])], [m for m in ['January', 'February', 'March', 'April', 'May', 'June'] if (m in x[3])]
        x[2], x[3] = x[2][0] if x[2] else "None", x[3][0] if x[3] else "None"
    
    # Create dictionaries to match each name to its index
    attendance_dict = {ccc_attendance_list[i][0]: i for i in range(
        len(ccc_attendance_list))}
    last_yr_dict = {last_year[i][0]: i for i in range(len(last_year))}

    # Begin database creation
    ccc_database = []

    for person in requests:

        # If the person isn't in attendance, they don't exist, else we get their attendance
        if person[0] not in attendance_dict:
            att = [person[0], "Does not exist"]
        else:
            att = ccc_attendance_list[attendance_dict[person[0]]]
        
        # If the person didn't perform last year, we record their month and time as None, else we get their last year records
        last_yr = last_year[last_yr_dict[person[0]]] if person[0] in last_yr_dict else [person[0], "None", "None"]

        # Add up their request, attendance, and last year
        ccc_database.append(person + att + last_yr)

    # Cut out redundant data (their names are repeated thrice)
    ccc_database = [x[0:5] + x[6:7] + x[8:] for x in ccc_database]

    # Get data frame
    database_df = pandas.DataFrame(ccc_database, columns = ["Name: ", "Teacher: ", "First choice: ", "Second choice: ", "Requested time: ", "Attendance: ", "Month performed last year: ", "Last duration: "])

    # Write to output
    xlwriter = pandas.ExcelWriter(
        'Database.xlsx', engine='xlsxwriter')
    database_df.to_excel(xlwriter, sheet_name='Records')
    worksheet1 = xlwriter.sheets['Records']
    worksheet1.set_column(1, 10, 30)
    xlwriter.save()

# Get attendance, then flatten the dataset
attendance('CCC Attendance 2020.xlsx')
flatten_dataset()
