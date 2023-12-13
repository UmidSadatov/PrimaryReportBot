import asyncio, datetime, os
from aiogram import Bot, Dispatcher
from aiogram.types import Message, File, \
    ReplyKeyboardMarkup, ReplyKeyboardRemove, \
    KeyboardButton, FSInputFile
from aiogram.filters import Command
import reports

TOKEN = "6948878045:AAEwtFFuklyf-uEmDh2CWTyRgollbIWtPNE"


class User:
    def __init__(self, user_id):
        self.user_id = user_id
        self.mode = 'sending_files'
        self.file_list = []
        # self.files_id = []
        self.none_sample_files = []
        self.report_file = None
        self.mode = 'user'

    def add_new_file(self, new_file_name):
        self.file_list.append(new_file_name)

    def add_unsupported_file(self, original_name):
        self.none_sample_files_original_names.append(original_name)

    def clear_file_list(self):
        self.file_list = []
        self.none_sample_files_original_names = []


Users = {
    70270879: User(70270879),
    3882812: User(3882812)
}

ADMINS_ID = [
    70270879,  # Umid
    3882812    # Sobir aka
]


def get_allowed_users_list():
    with open('allowed_users.txt', 'r') as users_list:
        return [username.strip() for username in users_list]


def add_allowed_user(username: str):
    if username not in get_allowed_users_list():

        with open('allowed_users.txt', 'a') as users_list:
            users_list.write('\n' + username)

        return f"Новый пользователь добавлен"

    else:
        return f"Пользователь уже в списке разрешенных"


def remove_allowed_user(removed_username: str):
    if removed_username in get_allowed_users_list():
        newlist = [
            username for username in get_allowed_users_list()
            if username != removed_username
        ]
        with open('allowed_users.txt', 'w') as users_list:
            users_list.write('\n'.join(newlist))
        return f"Пользователь удален из списка разрешенных"

    else:
        return f"Пользователя нет в списке разрешенных"


async def get_start(message: Message, bot: Bot):
    if message.from_user.username in get_allowed_users_list():
        Users[message.from_user.id] = User(message.from_user.id)
        button = KeyboardButton(text="Сделать отчет")
        markup = ReplyKeyboardMarkup(keyboard=[[button]])
        await bot.send_message(message.from_user.id,
                               "Отправьте файлы для обработки",
                               reply_markup=markup
                               )
    else:
        await bot.send_message(message.from_user.id,
                               f"Нет разрешения...\n "
                               f"Ваш юзернейм: {message.from_user.username}",
                               reply_markup=ReplyKeyboardRemove()
                               )


async def add_user(message: Message, bot: Bot):
    if message.from_user.id in ADMINS_ID:
        Users[message.from_user.id].mode = 'adding_user'
        await bot.send_message(
            message.from_user.id,
            'Юзернейм пользователя, которого нужно добавить:'
        )


async def remove_user(message: Message, bot: Bot):
    if message.from_user.id in ADMINS_ID:
        Users[message.from_user.id].mode = 'removing_user'
        await bot.send_message(
            message.from_user.id,
            'Юзернейм пользователя, которого нужно удалить:'
        )


async def getting_file(message: File, bot: Bot):
    if message.from_user.username in get_allowed_users_list():
        if message.document:
            new_file_id = message.document.file_id
            print(new_file_id)
            file = await bot.get_file(new_file_id)
            file_path = file.file_path
            file_dir = f"excel_files/{str(message.from_user.id)}"
            original_name = message.document.file_name

            file_extension = original_name.split('.')[-1].lower()
            if file_extension == 'xls':
                new_file_name = datetime.datetime.now().strftime(
                    f"{original_name.split('.')[0]}_%d%m%Y%H%M%S%f.xls")
            elif file_extension == 'xlsx':
                new_file_name = datetime.datetime.now().strftime(
                    f"{original_name.split('.')[0]}_%d%m%Y%H%M%S%f.xlsx")
            else:
                new_file_name = original_name

            if not os.path.exists(file_dir):
                os.makedirs(file_dir)
            await bot.download_file(file_path,
                                    f"{file_dir}/{new_file_name}")
            Users[message.from_user.id].add_new_file(
                f"{file_dir}/{new_file_name}")

        elif message.text:
            if message.text == "Сделать отчет":
                # Reports = [reports.OriginalReport(filename) for
                #          filename in Users[message.from_user.id].file_list]
                new_reports_file_name = \
                    f"excel_files/{message.from_user.id}_reports.xlsx"
                exception_names, none_sample_files = \
                    reports.make_report_file(
                        *Users[message.from_user.id].file_list,
                        new_file_name=new_reports_file_name)

                my_file = FSInputFile(new_reports_file_name)

                if len(none_sample_files) > 0:
                    none_sample_file_names = \
                        [none_sample_file.split('/')[-1]
                         for none_sample_file
                         in none_sample_files]

                    await bot.send_message(
                        message.from_user.id,
                        f"Для следующих файлов шаблон не определен:\n" +
                        '\n'.join(none_sample_file_names)
                    )

                if (len(Users[message.from_user.id].file_list) >
                        len(none_sample_files)):
                    await bot.send_document(
                        message.from_user.id, my_file
                    )

                exception_names = list(set(exception_names))
                if 0 in exception_names:
                    exception_names.remove(0)

                if len(exception_names) > 0:
                    await bot.send_message(
                        message.from_user.id,
                        "В базе данных следующие препараты не найдены:\n" +
                        '\n'.join([str(name) for name in exception_names]) +
                        f'\n<i>(всего {len(exception_names)})</i>' +
                        '\n<i>*препараты не добавлены в отчет</i>',
                        parse_mode='HTML'
                    )

                Users[message.from_user.id].clear_file_list()

    elif message.from_user.id in ADMINS_ID and \
            Users[message.from_user.id].mode != 'user':

        if Users[message.from_user.id].mode == 'adding_user':
            await bot.send_message(
                message.from_user.id,
                add_allowed_user(message.text) +
                '\n Далее:'
            )

        elif Users[message.from_user.id].mode == 'removing_user':
            await bot.send_message(
                message.from_user.id,
                remove_allowed_user(message.text) +
                '\n Далее:'
            )


    else:
        await bot.send_message(
            message.from_user.id,
            f"Нет разрешения...\n "
            f"Ваш юзернейм: {message.from_user.username}",
            reply_markup=ReplyKeyboardRemove()
        )


async def main():
    bot = Bot(token=TOKEN)
    dp = Dispatcher()

    dp.message.register(get_start, Command('start', 'apply'))
    dp.message.register(add_user, Command('add'))
    dp.message.register(remove_user, Command('remove'))
    dp.message.register(getting_file)
    try:
        await dp.start_polling(bot)
    finally:
        await bot.session.close()


if __name__ == "__main__":
    asyncio.run(main())
