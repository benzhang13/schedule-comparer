import discord
from app.process_excel import receive_file, read_all_files, delete_file
from datetime import datetime
import os
dir_path = os.path.dirname(os.path.realpath(__file__))

client = discord.Client()


def run_bot():
    token = os.environ['BOT_KEY']
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
        if message.attachments is None:
            await message.channel.send(message.author.mention + ', you have to attach your timetable.'
                                                                ' You can find the template with $template')
        else:
            already_exists = receive_file(message.attachments[0].url, str(message.author.id))
            if already_exists:
                await message.channel.send(message.author.mention +
                                           ', you already have a submitted timetable.'
                                           ' Please use $delete if you would like to replace your timetable.')
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

            if len(available) == 0:
                await message.channel.send(message.author.mention + ', no one is currently available')
            else:
                guild = message.guild
                for user_id in available:
                    u = guild.get_member(user_id)
                    await message.channel.send(u.mention + ' ')
                await message.channel.send('have no classes at the moment')
    elif message.content.startswith('$delete'):
        already_exists = delete_file(str(message.author.id))
        if not already_exists:
            await message.channel.send(message.author.mention +
                                       ', you have no timetable to delete. Use $submit to submit one.')
        else:
            await message.channel.send(message.author.mention + '\'s timetable deleted')
    elif message.content.startswith('$template'):
        user = message.author
        file = discord.File(fp=dir_path + '/app/generic_sheet/generic_sheet.xlsx')
        await message.channel.send(content=user.mention + ', here you go!', file=file)
    elif message.content.startswith('$help'):
        await message.channel.send(message.author.mention +
            """
            Commands include:
            $template: gets generic excel template
            $submit: submits the attached spreadsheet as a timetable
            $delete: deletes your current timetable
            $free: mention everyone who has no classes at the moment
            """)
    elif message.content.startswith('$'):
        await message.channel.send('I don\'t recognize that command. Try $help for a list of commands.')


if __name__ == '__main__':
    run_bot()
