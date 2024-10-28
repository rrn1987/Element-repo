class Person:
    def __init__(self, fname, lname, age):
        self.fname = fname
        self.lname = lname
        self.age = age

    def print(self):
        print('Name: {} {} and Age : {}'.format(self.fname, self.lname, self.age))


P1 = Person('Raj', 'Kap', 20)
P2 = Person('Rahul', 'Nair', 37)
P1.print()
P2.print()
