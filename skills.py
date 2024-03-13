import pandas as pd
import sys

# set excel file var
file=sys.argv[1] if len(sys.argv) > 1 else "CME Skills.xlsx"
outputfile1=sys.argv[2] if len(sys.argv) > 1 else "new_dataframe.xlsx"

# Load the original Excel file
df = pd.read_excel(file)

# print(skills_sheet)

def parse_skills(skills_str):
    if isinstance(skills_str, str):
        skills = skills_str.split('|')
        parsed_skills = []
        for skill in skills:
            skill_number, _, skill_info = skill.partition(' - ')
            if skill_info: # Check if there's a skill info after the partition
                skill_name, _, skill_level = skill_info.partition(' (')
                skill_level = skill_level.strip(')')
                if 'Advanced' in skill_level:
                    parsed_skills.append((skill_name, 'A'))
                elif 'Intermediate' in skill_level:
                    parsed_skills.append((skill_name, 'I'))
                elif 'Beginner' in skill_level:
                    parsed_skills.append((skill_name, 'B'))
                elif 'Master' in skill_level:
                    parsed_skills.append((skill_name, 'M'))
                elif 'Expert' in skill_level:
                    parsed_skills.append((skill_name, 'E'))
            else:
                # Handle cases where the delimiter is not found
                print(f"Unexpected format in skill entry: {skill}")
        return parsed_skills
    else:
        # Handle non-string values (e.g., NaN)
        return []


# Apply the parsing function to each row
df['ParsedSkills'] = df['Resource Skills'].apply(parse_skills)

# Flatten the list of skills and levels into separate columns
df_expanded = df.explode('ParsedSkills')
df_expanded[['Skill', 'Level']] = pd.DataFrame(df_expanded['ParsedSkills'].tolist(), index=df_expanded.index)

# Drop the original skills column
df_expanded.drop(columns=['Resource Skills', 'ParsedSkills'], inplace=True)

# write to new excel file
df_expanded.to_excel(outputfile1, sheet_name='test sheet', index=False)

