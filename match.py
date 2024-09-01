import pandas as pd
import random

mentors_df = pd.read_excel('mentors.xlsx')
mentees_df = pd.read_excel('mentees.xlsx')

def find_best_match(mentee, available_mentors):
    for criterion in ['major', 'class', 'faculty', 'hobbies']:
        if criterion in mentee:  
            matching_mentors = [mentor for mentor in available_mentors if pd.notna(mentee[criterion]) and mentee[criterion] == mentor[criterion]]
            if matching_mentors:
                return random.choice(matching_mentors)  
    return random.choice(available_mentors)

matches = []
mentor_assignments = {mentor['name']: 0 for mentor in mentors_df.to_dict('records')}

available_mentors = mentors_df.to_dict('records')
mentees_to_match = mentees_df.to_dict('records')

for mentor in available_mentors:
    best_match = find_best_match(mentor, mentees_to_match)
    if best_match:  
        
        matches.append({'Mentee': best_match['name'], 'Mentor': mentor['name']})
        mentor_assignments[mentor['name']] += 1
        mentees_to_match = [mentee for mentee in mentees_to_match if mentee['name'] != best_match['name']]
        if mentor_assignments[mentor['name']] == 2:
            available_mentors.remove(mentor)

for mentee in mentees_to_match:
    available_mentors = [mentor for mentor in mentors_df.to_dict('records') if mentor_assignments[mentor['name']] == 1]
    if available_mentors:
        best_match = find_best_match(mentee, available_mentors)
        matches.append({'Mentee': mentee['name'], 'Mentor': best_match['name']})
        mentor_assignments[best_match['name']] += 1

remaining_mentees = [mentee for mentee in mentees_df.to_dict('records') if not any(match['Mentee'] == mentee['name'] for match in matches)]

for mentee in remaining_mentees:
    available_mentors = [mentor for mentor in mentors_df.to_dict('records') if mentor_assignments[mentor['name']] < 2]
    if available_mentors:
        best_match = find_best_match(mentee, available_mentors)
        matches.append({'Mentee': mentee['name'], 'Mentor': best_match['name']})
        mentor_assignments[best_match['name']] += 1

matches_df = pd.DataFrame(matches)

matches_df.to_excel('mentee_mentor_matches.xlsx', index=False)

print("Matching completed and saved to mentee_mentor_matches.xlsx")
