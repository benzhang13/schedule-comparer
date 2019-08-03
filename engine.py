import discord
from app.process_excel import receive_file, read_all_files, delete_file
from app.process_ical import receive_file as ical_receive_file
from datetime import datetime
from dotenv import load_dotenv
import os
dir_path = os.path.dirname(os.path.realpath(__file__))

client = discord.Client()
load_dotenv()


def run_bot():
    token = os.environ['BOT_TOKEN']
    client.run(token)


@client.event
async def on_ready():
    print('{0.user} about to compare timetables'.format(client))


@client.event
async def on_message(message):
    if message.author == client.user:
        return

    if message.content.startswith('$hello'):
        await message.channel.send(message.author.mention + ', hey!')
    elif message.content.startswith('$submit'):
        if len(message.attachments) == 0:
            await message.channel.send(message.author.mention + ', you have to attach your timetable.'
                                                                ' You can find the template with $template')
        else:
            result = receive_file(message.attachments[0].url, str(message.author.id))
            if result.get('already_exists'):
                await message.channel.send(message.author.mention +
                                           ', you already have a submitted timetable.'
                                           ' Please use $delete if you would like to replace your timetable.')
            elif result.get('not_xlsx'):
                await message.channel.send(message.author.mention +
                                           ', that file is not a .xlsx file.'
                                           ' Please use the excel template provided with the $template command')
            else:
                await message.channel.send(message.author.mention + '\'s timetable submitted')
    elif message.content.startswith('$free'):
        now = datetime.today().timetuple()
        semester = 0
        if now[1] > 9 and now[2] >= 5:
            semester = 1
        elif now[1] <= 4 and now[2] >= 6:
            semester = 2

        if now[3] < 9 or now[3] > 21:
            await message.channel.send('No one has classes right now')
        else:
            if now[4] < 30:
                minute = 0
            else:
                minute = 30
            available = read_all_files(minute=minute, hour=now[3], day=datetime.today().weekday(), semester=semester)

            available_in_guild = []
            for user_id in available:
                guild = message.guild
                u = guild.get_member(user_id)
                if u is not None:
                    available_in_guild.append(u)
            if len(available_in_guild) == 0:
                await message.channel.send(message.author.mention + ', no one is currently available')
            else:
                for user in available_in_guild:
                    await message.channel.send(user.mention + ' ')
                await message.channel.send('have no classes at the moment')
    elif message.content.startswith('$delete'):
        already_exists = delete_file(str(message.author.id))
        if not already_exists:
            await message.channel.send(message.author.mention +
                                       ', you have no timetable to delete. Use $submit to submit one.')
        else:
            await message.channel.send(message.author.mention + '\'s timetable deleted')
    elif message.content.startswith('$get'):
        file_path = dir_path + '/app/spreadsheets/' + str(message.author.id) + 'schedule.xlsx'
        if os.path.isfile(file_path):
            file = discord.File(fp=file_path, filename=message.author.name + '_schedule.xlsx')
            await message.channel.send(content=message.author.mention+', here you go!', file=file)
        else:
            await message.channel.send('It appears that you do not currently have a submitted schedule.'
                                       ' Please use $help for commands to make one')
    elif message.content.startswith('$template'):
        user = message.author
        file = discord.File(fp=dir_path + '/app/generic_sheet/generic_sheet.xlsx')
        await message.channel.send(content=user.mention + ', here you go!', file=file)
    elif message.content.startswith('$subcal summer'):
        if len(message.attachments) == 0:
            await message.channel.send(message.author.mention + ', you must include the ICalendar file downloaded'
                                                                ' from the ACORN website as an attachment.')
        else:
            results = ical_receive_file(message.attachments[0].url, message.author.id, 0)

            if results.get('not_ics'):
                await message.channel.send(message.author.mention + ', that is not a .ics file. Make sure you are'
                                                                    ' submitting the file downloaded from the ACORN'
                                                                    ' website.')
            if results.get('submitted'):
                await message.channel.send(message.author.mention + ', your timetable has successfully been edited.')
    elif message.content.startswith('$subcal fall'):
        if len(message.attachments) == 0:
            await message.channel.send(message.author.mention + ', you must include the ICalendar file downloaded'
                                                                ' from the ACORN website as an attachment.')
        else:
            results = ical_receive_file(message.attachments[0].url, message.author.id, 1)

            if results.get('not_ics'):
                await message.channel.send(message.author.mention + ', that is not a .ics file. Make sure you are'
                                                                    ' submitting the file downloaded from the ACORN'
                                                                    ' website.')
            if results.get('submitted'):
                await message.channel.send(message.author.mention + ', your timetable has successfully been edited.')
    elif message.content.startswith('$subcal winter'):
        if len(message.attachments) == 0:
            await message.channel.send(message.author.mention + ', you must include the ICalendar file downloaded'
                                                                ' from the ACORN website as an attachment.')
        else:
            results = ical_receive_file(message.attachments[0].url, message.author.id, 2)

            if results.get('not_ics'):
                await message.channel.send(message.author.mention + ', that is not a .ics file. Make sure you are'
                                                                    ' submitting the file downloaded from the ACORN'
                                                                    ' website.')
            if results.get('submitted'):
                await message.channel.send(message.author.mention + ', your timetable has successfully been edited.')
    elif message.content.startswith('$help'):
        await message.channel.send(message.author.mention +
            """
            Commands include:
            $template: gets generic excel template
            $submit: submits the attached spreadsheet as a timetable
            $delete: deletes your current timetable
            $get: gets your current submitted schedule as an excel file
            $free: mention everyone who has no classes at the moment
            $subcal: edits timetable using ical file from ACORN. Use <$subcal help> for more details about this
            """)
    elif message.content.startswith('$subcal help'):
        await message.channel.send(message.author.mention +
            """
            The $subcal command is used for importing ical(.ics) schedules downloaded from ACORN.
            To find how to retrieve your ical schedule from the ACORN site, use <$subcal retrieve help>.
            Once you have your ical schedule for the semester, use the command:
            $subcal <semester>, where <semester> semester is the semester for your ical schedule, i.e. summer, fall, winter.
            Make sure the ical file is attached to your command message.
            If you do not currently have a timetable, this command will create one for you.
            """)
    elif message.content.startswith('$subcal retrieve help'):
        await message.channel.send(message.author.mention +
            """
            Steps to retrieve your ical schedule from ACORN:
            1. Log in to ACORN
            2. Click the "View Timetable" button on the home page
            3. Click "Download Calendar Export"
            4. Your ical file for the current semester should begin downloading
            """)
    elif message.content.startswith('$'):
        await message.channel.send('I don\'t recognize that command. Try $help for a list of commands.')


if __name__ == '__main__':
    run_bot()
