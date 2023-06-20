import discord
from discord.ext import commands, tasks
import pandas as pd
import openpyxl
import datetime

# Set up the bot client
intents = discord.Intents.all()
intents.members = True  # Enable the Members intent
intents.message_content = True
intents.presences = True


bot = commands.Bot(command_prefix='!', intents=intents)

# Set up the path to the spreadsheet
SPREADSHEET_PATH = 'clock_data.xlsx'

# Load the spreadsheet (create a new one if it doesn't exist)
try:
    df = pd.read_excel(SPREADSHEET_PATH, engine='openpyxl')
except FileNotFoundError:
    df = pd.DataFrame()  # Create an empty DataFrame

@bot.event
async def on_ready():
    print('Bot is ready.')
    check_clocked_in_users.start()

@tasks.loop(minutes=15)
async def check_clocked_in_users():
    try:
        now = datetime.datetime.now()
        date = now.date()
        time = now.strftime('%H:%M')  # Format time as 'hour:minute'
        df = pd.read_excel(SPREADSHEET_PATH, engine='openpyxl')

        # Check if 'User ID' column exists in the spreadsheet
        if 'User ID' not in df.columns:
            df['User ID'] = ''

        # Get the guild object
        guild = bot.get_guild(1119600234620788789)  # Replace with your actual guild ID

        if guild is None:
            print("Guild not found.")
            return

        for _, row in df.iterrows():
            user_id = row['User ID']
            username = row['Username']
            clock_in_date = row['Clock-in Date']
            clock_in_time = row['Clock-in Time']

        # Skip users who are already clocked out or have no clock-in data
        if pd.isnull(row['Clock-out Date']) and clock_in_date is not None and clock_in_time is not None:
            member = await guild.fetch_member(user_id)

        # Check if the member is offline
        if member is not None and member.status == discord.Status.offline:
            # Clock out the user
            df.loc[df['User ID'] == user_id, 'Clock-out Date'] = date
            df.loc[df['User ID'] == user_id, 'Clock-out Time'] = time

            # Calculate the working hours
            working_hours = calculate_working_hours(clock_in_time, time)

            # Add the working hours to the corresponding row
            df.loc[df['User ID'] == user_id, 'Working Hours'] = working_hours
            df.to_excel(SPREADSHEET_PATH, index=False)

            await bot.get_channel(1119600235119923203).send(f'{username} has been automatically clocked out due to being offline.')


    except Exception as e:
        print(f'An error occurred during check_clocked_in_users task: {e}')
    df.to_excel(SPREADSHEET_PATH, index=False)

@bot.command()
async def clockin(ctx):
    try:
        author_id = ctx.author.id
        author_name = ctx.author.name
        now = datetime.datetime.now()
        date = now.date()
        time = now.strftime('%H:%M')  # Format time as 'hour:minute'
        df = pd.read_excel(SPREADSHEET_PATH, engine='openpyxl')

        # Check if 'User ID' column exists in the spreadsheet
        if 'User ID' not in df.columns:
            df['User ID'] = ''

        # Check if the user already has a clock-in entry
        existing_row = df[(df['User ID'] == author_id)]  # Convert author_id to string for comparison
        if not existing_row.empty:
            # Check if the previous clock-out data is present
            if pd.isnull(existing_row.iloc[-1]['Clock-out Date']):
                await ctx.send(f'{ctx.author.mention} has not clocked out from the previous shift.')
                return

        # Update the existing row with the clock-in timestamp or create a new row
        if not existing_row.empty:
            df.loc[existing_row.index, 'Clock-in Date'] = date
            df.loc[existing_row.index, 'Clock-in Time'] = time
        else:
            new_data = {'User ID': str(author_id), 'Username': author_name, 'Clock-in Date': date, 'Clock-in Time': time, 'Clock-out Date': '', 'Clock-out Time': ''}
            df = pd.concat([df, pd.DataFrame(new_data, index=[0])], ignore_index=True)

        df.to_excel(SPREADSHEET_PATH, index=False)
        await ctx.send(f'{ctx.author.mention} has clocked in at {date} {time}.')

    except Exception as e:
        print(f'An error occurred during clock_in command: {e}')
        await ctx.send('An error occurred while processing the command. Please try again later.')

@bot.command()
async def clockout(ctx):
    try:
        author_id = ctx.author.id
        author_name = ctx.author.name
        now = datetime.datetime.now()
        date = now.date()
        time = now.strftime('%H:%M')  # Format time as 'hour:minute'
        df = pd.read_excel(SPREADSHEET_PATH, engine='openpyxl')

        # Check if 'User ID' column exists in the spreadsheet
        if 'User ID' not in df.columns:
            df['User ID'] = ''

        # Check if the user already has a clock-in entry
        existing_row = df[df['User ID'] == author_id]
        if not existing_row.empty:
            # Check if the previous clock-out data is already present
            if pd.isnull(existing_row.iloc[-1]['Clock-out Date']):
                # Update the existing row with the clock-out timestamp
                df.loc[existing_row.index, 'Clock-out Date'] = date
                df.loc[existing_row.index, 'Clock-out Time'] = time

                # Calculate the working hours
                clock_in_time = existing_row.iloc[-1]['Clock-in Time']
                clock_out_time = time
                working_hours = calculate_working_hours(clock_in_time, clock_out_time)

                # Add the working hours to the corresponding row
                df.loc[existing_row.index, 'Working Hours'] = working_hours
            else:
                await ctx.send(f'{ctx.author.mention} has already clocked out.')
                return
        else:
            await ctx.send(f'{ctx.author.mention} has not clocked in.')

        df.to_excel(SPREADSHEET_PATH, index=False)
        await ctx.send(f'{ctx.author.mention} has clocked out at {date} {time}.')

    except Exception as e:
        print(f'An error occurred during clockout command: {e}')
        await ctx.send('An error occurred while processing the command. Please try again later.')

def calculate_working_hours(clock_in_time, clock_out_time):
    fmt = '%H:%M'
    clock_in = datetime.datetime.strptime(clock_in_time, fmt)
    clock_out = datetime.datetime.strptime(clock_out_time, fmt)
    working_hours = clock_out - clock_in
    return str(working_hours)

@bot.command()
@commands.has_role('admin')  # Check if the user has the 'admin' role
async def viewlog(ctx):
    try:
        admin_user = ctx.author
        df = pd.read_excel(SPREADSHEET_PATH, engine='openpyxl')
        if 'Username' not in df.columns or 'Working Hours' not in df.columns:
            await admin_user.send('The log does not exist or does not contain the required columns.')
            return

        log_data = df[['Username', 'Working Hours']].copy()
        log_data = log_data.groupby('Username').sum()

        if log_data.empty:
            await admin_user.send('No log data available.')
            return

        log_data = log_data.reset_index()
        log_data = log_data.rename(columns={'Working Hours': 'Total Working Hours'})

        log_data_str = log_data.to_string(index=False)
        await admin_user.send(f'```\n{log_data_str}\n```')
    except Exception as e:
        print(f'An error occurred during viewlog command: {e}')
        await admin_user.send('An error occurred while processing the command. Please try again later.')

bot.run('MTExOTYwMTE0MjIxNzgzODY5Mw.GOD58y.RQs0b9ltGsrkPkDaYFWGjuFgSkwma4JSTSTQMU')