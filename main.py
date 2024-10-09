import os
import openai
from openai import OpenAI
import dotenv
import pandas as pd


dotenv.load_dotenv()

openai_api_key = os.getenv("OPENAI_API_KEY")
client = openai.Client(api_key=openai_api_key)
MODEL = "gpt-4o"


def remove_first_person_from_descriptions(text):
    try:
        messages=[
            # change the prompt here
            {"role": "system", "content": "Convert the following product description from first person to third person, replacing all references to 'we', 'our', 'us', or similar with the actual vendor's name or other relevant terms that make sense in context. Only make the necessary adjustments to fix grammar while maintaining the original meaning. Do not change anything else in the description. Additionally, replace any instances of '<p><strong>' tags with '<h4>' tags and Remove any other HTML styling, including inline styles or attributes, from inside existing tags. Ensure that '<h4>' tags do not contain colons (:) in their text. Ensure that the output format remains in valid HTML. Keep '<p>' tags as they are unless they are '<p><strong>' tags, in which case replace them with '<h4>' tags. Ensure that '<h4>' tags do not contain colons (:) in their text. Ensure that '<h4>' tags do not contain colons (:) in their text. Remove any other HTML styling, including inline styles or attributes, from inside existing tags. Ensure that the output format remains in valid HTML. Replace bullet lists using '•' with HTML '<ul>' and '<li>' tags. Remove any contact information or text suggesting making contact, such as phone numbers, email addresses, or phrases like 'contact us' or 'get in touch'."},
            {"role": "user", "content": text}
        ]
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            temperature=0.0,
        )
        gpt_response = response.choices[0].message.content
        return gpt_response
    except Exception as e:
        print("Error in remove_first_person_from_text", e)
        return None
    

def remove_first_person_from_features(text):
    try:
        messages=[
            # change the prompt here
            {"role": "system", "content": "Review the following features list and make minimal adjustments to ensure it is in a clean HTML list format using '<ul>' and '<li>' tags. If the list is already in good format, leave it as is. If there are any breaks in the list, such as separate lists or inconsistencies, merge them into a single unified list. Remove any inline styles or attributes from the HTML tags.Also, replace any first-person references like 'we', 'our', 'us' with the relevant third-person terms that make sense in context. Ensure that the output remains in valid HTML."},
            {"role": "user", "content": text}
        ]
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            temperature=0.0,
        )
        gpt_response = response.choices[0].message.content
        return gpt_response
    except Exception as e:
        print("Error in remove_first_person_from_text", e)
        return None


if __name__ == "__main__":
    df = pd.read_excel("First person fix.xlsx")
    df["New Product Description"] = None
    df["New Features"] = None
    for index, row in df.iterrows():
        print(f"Processing row {index}")
        product_description = row["Product Description"]
        feature = row["Features"]
        if pd.isna(product_description) and pd.isna(feature):
            print(f"Row {index} has both cells empty, skipping")
            continue
        if not pd.isna(product_description):
            new_description = remove_first_person_from_descriptions(text=product_description)
            df.at[index, "New Product Description"] = new_description
        if not pd.isna(feature):
            new_feature = remove_first_person_from_features(text=feature)
            df.at[index, "New Features"] = new_feature
        df.to_excel("First person fix-updated.xlsx", index=False, engine='openpyxl')
        print(f"Row {index} complete and saved to file")

    print("All rows processed")
    

