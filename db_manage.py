import sqlite3

db_con = sqlite3.connect('reports.db')
db_con.row_factory = sqlite3.Row
cursor = db_con.cursor()


# def insert_unique_name(unique_name, general_id):
#     cursor.execute("""INSERT INTO All_names
#     (unique_name, general_id) VALUES (?,?)""",
#                    (unique_name, general_id,))
#     db_con.commit()


def get_general_name_and_price(unique_name):
    while unique_name[-1] == ' ':
        unique_name = unique_name[:-1]

    cursor.execute(
        f"""SELECT name, price 
        FROM General_Names INNER JOIN All_names 
        ON General_Names.id = All_names.general_id 
        WHERE unique_name = '{unique_name}'"""
    )
    result = cursor.fetchone()
    db_con.commit()
    return result[0], float(result[1].replace(',', '.'))


def get_group(general_name):
    cursor.execute(
        f"""SELECT "group"
        FROM General_Names INNER JOIN Groups
        ON General_Names.group_id = Groups.group_id
        WHERE name='{general_name}'"""
    )
    group = cursor.fetchone()
    db_con.commit()
    return group[0]


# name, price = get_general_name_and_price("ПРОЖЕСТОЖЕЛЬ ГЕЛЬ 1% 80 Г")

# print(get_general_name('asdads'))


# print(get_general_name("Таблетка Гексализ-30"))

# cursor.execute("""
# UPDATE General_Names
# SET price = REPLACE(price, '$', '')
# WHERE price LIKE '$%';
# """)
#
# cursor.execute("""
# UPDATE General_Names
# SET price = REPLACE(price, '€ ', '')
# WHERE price LIKE '€%';
# """)

# db_con.commit()
