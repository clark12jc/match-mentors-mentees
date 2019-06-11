import pandas as pd
import Person


# Not using this - was going to be used to split markets into a list
# Use format_market instead
def split_markets(markets):
    if markets.find(';') != -1:
        markets = markets.split(';')
        for market in markets:
            market.strip()
    return markets


def format_market(market):
    if market.find(';') != -1:
        semi = market.find(';')
        market = market[:semi]
    return market


def normalize_location(location):
    dc_strings = ['dc', 'metro', 'arlington']
    if any(dc_string in location.lower() for dc_string in dc_strings):
        return 'Washington, DC'
    else:
        return location


def format_name(name):
    comma = name.find(',')
    if comma != -1:
        last_name, first_name = name.split(', ', 1)
    else:
        first_name, last_name = name.split(' ', 1)
    return first_name, last_name


def match_one(one, two):
    return (
        one.location.lower() == two.location.lower() and
        one.capability.lower() == two.capability.lower() and
        one.market.lower() == two.market.lower())


def match_two(one, two):
    return (
        one.location.lower() == two.location.lower() and
        two.capability.lower() == two.capability.lower()
    )


def match_three(one, two):
    return one.location.lower() == two.location.lower()


# Read mentor list excel file
def read_excel_mentor_list():
    df = pd.read_excel("mentor-data.xlsx", skipinitialspace=True, dtype=str)

    # Strip whitespace & replace 'nan' values
    df = df.replace(' ', '')
    df = df.replace('nan', '')
    df = df.fillna(value='')

    # Converts dataframe to dictionary list
    df_list = df.to_dict(orient='records')

    mentor_list = []
    for mentor in df_list:
        # Assign excel columns to variables
        first = mentor['First'].strip()
        last = mentor['Last'].strip()
        location = mentor['Location'].strip()
        capability = mentor['Capability'].strip()
        market = mentor['Market'].strip()
        market = format_market(market)
        email = mentor['Email'].strip()

        # Instantiate new Person
        person = Person.Person(first, last, location, capability, market, email, True)

        # Skip duplicate mentors
        if person not in mentor_list:
            mentor_list.append(person)
    return mentor_list


# Read mentee list excel file
def read_excel_mentee_list():
    df = pd.read_excel("mentee-data.xlsx", skipinitialspace=True, dtype=str)

    # Strip whitespace & replace 'nan' values
    df = df.replace(' ', '')
    df = df.replace('nan', '')
    df = df.fillna(value='')

    # Converts dataframe to dictionary list
    df_list = df.to_dict(orient='records')

    mentee_list = []

    for mentee in df_list:
        # Assign excel columns to variables
        first = mentee['First'].strip()
        last = mentee['Last'].strip()
        location = mentee['Location'].strip()
        capability = mentee['Capability'].strip()
        market = mentee['Market'].strip()
        market = format_market(market)
        email = mentee['Email'].strip()

        # Instantiate new Person
        person = Person.Person(first, last, location, capability, market, email, False)

        # Skip duplicate mentees
        if person not in mentee_list:
            mentee_list.append(person)
    return mentee_list


# Creates a structured dictionary format for easy output to excel
def format_match_for_excel(mentor, mentee):
    # TODO - Added emails here so they output to excel; left the original output commented out
    match_dict = {
        'Mentor': mentor.full_name,
        'Mentee': mentee.full_name,
        'Mentee Location': mentee.location
    }
    # match_dict = {
    #     'Mentor': mentor.full_name,
    #     'Mentee': mentee.full_name,
    #     'Mentee Email': mentee.email,
    #     'Mentee Location': mentee.location
    # }
    return match_dict


def write_matches_to_excel(match_list):
    new_list = []
    for match in match_list:
        match_dictionary = format_match_for_excel(match['Mentor'], match['Mentee'])
        new_list.append(match_dictionary)
    match_list = new_list

    try:
        df = pd.DataFrame(match_list)
        writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter', options={'strings_to_urls': False})
        df.to_excel(writer)
        writer.save()
    except Exception as e:
        print(e)
    return


def main():
    # Get mentor and mentee lists from excel
    mentor_list = read_excel_mentor_list()
    mentee_list = read_excel_mentee_list()
    match_list = []

    print('Number of Mentors:', str(len(mentor_list)))
    print('Number of Mentees:', str(len(mentee_list)))

    # Do match One
    for mentor in mentor_list:
        for mentee in mentee_list:
            # Skip matches
            if mentor.has_match:
                break
            if mentee.has_match:
                continue

            if match_one(mentor, mentee):
                match = {
                    'Mentor': mentor,
                    'Mentee': mentee,
                    'Level': 'one'
                }
                # Update has_match variable
                mentor.has_match = True
                mentee.has_match = True
                # Add to match list
                match_list.append(match)
                # mentor_list.remove(mentor)
                # mentee_list.remove(mentee)
                break

    # Do match Two
    for mentor in mentor_list:
        for mentee in mentee_list:
            # Skip matches
            if mentor.has_match:
                break
            if mentee.has_match:
                continue

            if match_two(mentor, mentee):
                match = {
                    'Mentor': mentor,
                    'Mentee': mentee,
                    'Level': 'two'
                }
                # Update has_match variable
                mentor.has_match = True
                mentee.has_match = True
                # Add to match list
                match_list.append(match)
                # mentor_list.remove(mentor)
                # mentee_list.remove(mentee)
                break

    # Do match Three
    for mentor in mentor_list:
        for mentee in mentee_list:
            # Skip matches
            if mentor.has_match:
                break
            if mentee.has_match:
                continue

            if match_three(mentor, mentee):
                match = {
                    'Mentor': mentor,
                    'Mentee': mentee,
                    'Level': 'three'
                }
                # Update has_match variable
                mentor.has_match = True
                mentee.has_match = True
                # Add to match list
                match_list.append(match)
                # mentor_list.remove(mentor)
                # mentee_list.remove(mentee)
                break

    # Match leftovers
    for mentor in mentor_list:
        for mentee in mentee_list:
            # Skip matches
            if mentor.has_match:
                break
            if mentee.has_match:
                continue

            match = {
                'Mentor': mentor,
                'Mentee': mentee,
                'Level': 'four'
            }
            # Update has_match variable
            mentor.has_match = True
            mentee.has_match = True
            # Add to match list
            match_list.append(match)
            # Do NOT remove item from mentor_list
            # mentee_list.remove(mentee)
            break

    print('Total Matches:', len(match_list))
    write_matches_to_excel(match_list)

    # List mentors without a match
    no_match_mentors = []
    for mentor in mentor_list:
        if mentor.has_match is False:
            no_match_mentors.append(mentor)


# Instantiate main 888888888888888888888888888888888888888888888888888888888888
if __name__ == "__main__":
    main()
