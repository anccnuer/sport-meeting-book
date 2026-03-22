class Student:
    def __init__(self, id=None, grade=None, class_name=None, gender=None, number=None, name=None):
        self.id = id
        self.grade = grade
        self.class_name = class_name
        self.gender = gender
        self.number = number
        self.name = name

class Event:
    def __init__(self, id=None, name=None):
        self.id = id
        self.name = name

class StudentEvent:
    def __init__(self, id=None, student_id=None, event_id=None):
        self.id = id
        self.student_id = student_id
        self.event_id = event_id
