import main
import pandas as pd


def test_unique_matches():
    print('\nTest - Unique People in Match')
    df = pd.read_excel("output.xlsx", skipinitialspace=True, dtype=str)

    # Strip whitespace & replace 'nan' values
    df = df.replace(' ', '')
    df = df.replace('nan', '')
    df = df.fillna(value='')
    df_list = df.to_dict(orient='records')

    mentor_list = []
    mentee_list = []
    for match in df_list:
        mentor = match['Mentor']
        mentee = match['Mentee']

        # Check for duplicate mentors
        if mentor in mentor_list:
            print('Duplicate Mentor:', mentor)
        else:
            mentor_list.append(mentor)

        # Check for duplicate mentees
        if mentee in mentee_list:
            print('Duplicate Mentee:', mentee)
        else:
            mentee_list.append(mentee)
    return


def test_normalize_location():
    print('\nTest - normalize_location')
    print(main.normalize_location('Washington DC area'))  # Washington, DC
    print(main.normalize_location('McLean, VA'))  # McLean, VA
    print(main.normalize_location('Arlington, VA'))  # Washington, DC


def test_split_markets():
    print('\nTest - split_markets')
    markets = 'Photos; Geography;Cybersecurity'
    print(main.split_markets(markets))  # ['Photos', ' Geography', 'Cybersecurity']

    markets = 'Financial Services; Energy, Resources, and Utilities'
    print(main.split_markets(markets))  # ['Financial Services', ' Energy, Resources, and Utilities']


def test_format_name():
    print('\nTest - format_name')
    print(main.format_name('Gordon, Josh'))  # ('Josh', 'Gordon')
    print(main.format_name('Josh Gordon'))  # ('Josh', 'Gordon')


if __name__ == '__main__':
    test_format_name()
    test_split_markets()
    test_normalize_location()
    test_unique_matches()
