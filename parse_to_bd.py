import psycopg2
import json


def get_weekday_number(day):
    weekdays = {
        'понедельник': 1,
        'вторник': 2,
        'среда': 3,
        'четверг': 4,
        'пятница': 5,
        'суббота': 6,
        'воскресенье': 7
    }
    return weekdays.get(day.lower(), None)


def get_week_number(week):
    weeks = {
        'I': 1,
        'II': 2
    }
    return weeks.get(week, None)


# Подключение к базе данных
conn = psycopg2.connect(
        dbname="public",
        user="usre",
        password="usre",
        host="localhost",
    )
cur = conn.cursor()

cur.execute("ALTER TABLE sc_rasp ALTER COLUMN subgroup SET DEFAULT 0;")


# Чтение JSON файла
for i in range(1, 7):
    with open(f'kurs_work-{i}.json', encoding="utf-8") as f:
        data = json.load(f)

    for group, schedule in data.items():
        print(group)
        print(schedule)
        for entry in schedule:
            print(entry)
            # Проверка наличия всех необходимых полей
            if all(key in entry for key in
                   ['День недели', 'Номер пары', 'Нач. занятия', 'Оконч. занятия', 'Неделя занятия', 'Дисциплина занятия',
                    'ФИО преподавателя']):
                # Добавление данных в sc_disc
                if entry['Дисциплина занятия'] != '':
                    cur.execute("INSERT INTO public.sc_disc (title) SELECT %s WHERE NOT EXISTS (SELECT 1 FROM public.sc_disc WHERE title = %s)",
                                (entry['Дисциплина занятия'], entry['Дисциплина занятия']))

                # Добавление данных в sc_group
                cur.execute(
                    "INSERT INTO public.sc_group (title) SELECT %s WHERE NOT EXISTS (SELECT 1 FROM public.sc_group WHERE title = %s)",
                    (group, group))

                # Добавление данных в sc_prep
                if entry['ФИО преподавателя'] != '':
                    cur.execute(
                        "INSERT INTO public.sc_prep (fio) SELECT %s WHERE NOT EXISTS (SELECT 1 FROM public.sc_prep WHERE fio = %s)",
                        (entry['ФИО преподавателя'], entry['ФИО преподавателя']))

                # Получение id добавленных записей
                cur.execute("SELECT id FROM public.sc_disc WHERE title = %s", (entry['Дисциплина занятия'],))
                row = cur.fetchone()
                if row is not None:
                    disc_id = row[0]
                else:
                    disc_id = None

                cur.execute("SELECT id FROM public.sc_group WHERE title = %s", (group,))
                group_id = cur.fetchone()[0]

                cur.execute("SELECT id FROM public.sc_prep WHERE fio = %s", (entry['ФИО преподавателя'],))
                row1 = cur.fetchone()
                if row1 is not None:
                    prep_id = row1[0]
                else:
                    prep_id = None

                # Добавление данных в sc_rasp с указанием subgroup = 0
                cur.execute(
                    "INSERT INTO public.sc_rasp (disc_id, prep_id, weekday, week, lesson, group_id, subgroup) VALUES (%s, %s, %s, %s, %s, %s, DEFAULT)",
                    (disc_id, prep_id, get_weekday_number(entry['День недели']),
                     get_week_number(entry['Неделя занятия']),
                     int(entry['Номер пары']), group_id))

# Фиксация изменений и закрытие соединения
conn.commit()
cur.close()
conn.close()
