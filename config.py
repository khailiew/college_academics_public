# Variables, constants, and Classes for college_academics.py

import re
#FORMAT
# colours for excel sheet tabs
college_colours = {
    'BASS' : "5370cf", 
    'BAXT' : "8FC4FF", 
    'COLH' : "B1A0C7", 
    'FTH' : "C4D79B", 
    'GOLD' : "DA9694", 
    'IH' : "BFBFBF", 
    'HALL' : "ffd045",
    'ALL' : "000000"
    }
    
college_names = {
    'BASS' : "Basser College", 
    'BAXT' : "Baxter College", 
    'COLH' : "Colombo House", 
    'FTH' : "Fig Tree Hall", 
    'GOLD' : "Golstein College", 
    'IH' : "International House", 
    'HALL' : "UNSW Hall",
    }

# excel column col_widths, add any additional columns in order
col_widths = {
    "First Names" : 18,
    "Last Name" : 14, 
    "zID" : 8, 
    "College" : 7,
    "Type" : 9,
    "Program" : 8,
    "Term" : 12,
    "Code" : 10,
    "Course": 40,
    "Mark" : 5,
    "Grade" : 16,
    "WAM" : 6
}

# Function to convert short code to term name or vice versa
def convert_term_name(term_string):
    term_string = term_string.strip()
    
    # return None if not matching term code pattern
    if re.match("(?i)^\d\d[ST][0123]$", term_string) is not None:
        term_code = term_string
        
        # get year, term/sem, and term/sem number
        year = term_code[:1]
        year = "20" + year
        
        if term_code[2].upper() == "S":
            term_sem = "SEMESTER"
        else:
            term_sem = "TERM"
        
        if term_code[-1] == "0":
            return f"{year} SUMMER {term_sem}"
        
        return f"{year} {term_sem} {term_code[-1]}"
    
    elif re.match("(?i)^\d{4} (?:TERM|SEMESTER) [0123]$", term_string) is not None or \
         re.match("(?i)^\d{4} SUMMER (?:TERM|SEMESTER)$", term_string) is not None:
        #convert to short code
        term_split = re.split(' ', term_string)
        if term_split[1].lower() == "summer":
            term_split[1] = "0"
            term_split.append(term_split.pop(1)) #move to end
            
        return term_split[0][-2:] + term_split[1][0].upper() + term_split[2]
    
    return None

#OBJECTS                
class Course:
    def __init__(self, code_name, code_num, name, grade='-', grade_name='-'):
        self.code = code_name + code_num
        self.name = name
        self.grade = grade
        self.grade_name = grade_name
    
    def hasGrade(self):
        return self.grade.isnumeric()
        
    def __repr__(self):
        return f'Course({self.code}, {self.name}, {self.grade}, {self.grade_name})'

class Student:
    def __init__(self, name, zid, college, enrol_type='', program=''):
        self.first_names = self.splitName(name)[0].title()
        self.last_name = self.splitName(name)[1]
        self.zid = zid
        self.college = college
        self.enrol_type = enrol_type.title() #UGRD/PGRD
        self.program = program
        self.terms = {}
        self.wams = {}
        self.overall_wam = None
        
    # dict to hold term courses
    def addCourse(self, term, course_obj):
        # initialise term if not exist yet
        if term not in self.terms:
            self.terms[term] = []
        self.terms[term].append(course_obj)
        
    # get first and last names
    def splitName(self, name):
        first_names = ' '.join(name.split()[:-1])
        last_name = name.split()[-1]
        return (first_names, last_name)
    
    #calculate and return WAM
    def calc_wam(self, select_term):
        if select_term in self.terms.keys():
            courses = self.terms[select_term]
            counter = 0
            total = 0
            for course in courses:
                if course.hasGrade():
                    counter += 1
                    total += float(course.grade)
            if counter:
                return f'{total/counter:.1f}'
            else:
                return None
    
    # calculate wams for all terms, and then get overall wam
    def process_wams(self):
        total = 0
        
        for t in self.terms:
            w = self.calc_wam(t)
            self.wams[t] = w
            if w is not None:
                total += float(w)
        
        if self.wams:           
            self.overall_wam = f'{total / sum(_ is not None for _ in self.wams):.1f}'
            
    def __repr__ (self):
        return f'Student( {self.first_names+" "+self.last_name}, {self.zid}, {self.enrol_type}, {self.program}, {self.wams}, {self.overall_wam}, \n{self.terms} )\n\n'
  