


students=[('Ania', 3.5), ('Kamil', 4.5), ('Ola', 4.0), ('Piotr', 5.0), ('Ewa', 3.0), ('Cezary', 2.4), ('Adam', 3.7), ('Barbara', 4.2), ('Daria', 4.9), ('Barbara', 3.2), ('Edward', 4.6), ('Cezary', 2.8), ('Adam', 4.4), ('Daria', 4.5), ('Barbara', 2.6), ('Edward', 3.6), ('Cezary', 4.7), ('Adam', 2.9), ('Daria', 3.5), ('Edward', 3.3), ('Cezary', 4.2), ('Daria', 4.2), ('Edward', 3.1), ('Barbara', 4.8), ('Edward', 4.8)]



students_above_4 = [student for student in students if student[1]>4]
students_above_4_dict = {student[0]:student[1] for student in students for student in students if student[1]>4}

# print(students_above_4)
print(students_above_4_dict)
