from peewee import *
from datetime import datetime
from config import MYSQL_CONN

db = MySQLDatabase(**MYSQL_CONN)


def before_request_handler():
    db.connect()


def after_request_handler():
    db.close()


class BaseModel(Model):
    class Meta:
        database = db


class Users(BaseModel):
    id = PrimaryKeyField()
    telegram_id = IntegerField(unique=True)
    username = CharField(default=None)
    name = CharField()
    dt = DateTimeField(default=datetime.now())


class Schedule(BaseModel):
    id = PrimaryKeyField()
    place = CharField(null=True)
    course_name = CharField(null=True)
    metro = CharField(null=True)
    address = CharField(null=True)
    comments = TextField(null=True)
    age_from = IntegerField(null=True)
    age_to = IntegerField(null=True)
    period = CharField(null=True)
    price = CharField(null=True)
    lecturer = CharField(null=True)
    day_of_week = CharField(null=True)
    time = CharField(null=True)
    upd_dt = DateTimeField(default=datetime.now())


class Tags(BaseModel):
    id = PrimaryKeyField()
    course_name = CharField(null=True)
    age_from = IntegerField(null=True)
    age_to = IntegerField(null=True)
    tag2 = CharField(null=True)
    tag3 = CharField(null=True)
    tag4 = CharField(null=True)
    upd_dt = DateTimeField(default=datetime.now())


def init_db():
    tables = [Schedule, Tags, Users]
    for t in tables:
        if t.table_exists():
            t.drop_table(cascade=True)
        t.create_table()


def save(data, db_name):
    with db.atomic():
        db_name.insert_many(data).upsert().execute()
    return True


if __name__ == '__main__':
    init_db()
    print('Таблицы создал')
