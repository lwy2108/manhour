class Employee:
  
  "An employee class to represent staff, their attributes and store leave entries."
  
  hours = "Regular"
  
  def __init__(self, name):
    self.name = name
  
  def check_shift(self, shift_work):
    if shift_work == 1:
      self.hours = "Shift"
    elif shift_work == 0:
      self.hours = hours
  
  def add_leave(self, leave):
    pass
