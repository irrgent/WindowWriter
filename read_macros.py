import csv


# Reads csv file line by line using it to create a dictionary of macros
def macro_dict(filePath):

    with open(filePath) as csvFile:

        # Reads with dialect='excel' by default
        csvreader = csv.reader(csvFile)
        csvdict = {}

        for row in csvreader:

            if len(row) < 2:
                raise ValueError("Row with less than 2 columns found.")
            elif len(row) > 2:
                raise ValueError("Row with more than 2 columns found.")

            csvdict[row[0]] = row[1]

    return csvdict


if __name__ == '__main__':

    d = macro_dict('macros.csv')
