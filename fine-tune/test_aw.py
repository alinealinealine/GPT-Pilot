# -*- coding: utf-8 -*-
"""
Created on Sun Mar  5 18:59:34 2023

@author: xweng

This file is to use API fine-tune the market part. 
Reference: 
    https://github.com/openai/openai-cookbook/blob/main/examples/Fine-tuned_classification.ipynb
    https://platform.openai.com/docs/guides/fine-tuning/preparing-your-dataset
    https://platform.openai.com/docs/api-reference/fine-tunes/create

"""
%env OPENAI_API_KEY = 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

import openai

import pandas as pd
import re

# Load data into a pandas dataframe
data = pd.read_csv('FIG_AIMM_text.csv', encoding='latin1').loc[:,['project_description', 'sector','country_description','text.summary']]

# Drop missing values (no IFC disclosure in project description)
data.dropna(inplace=True)


# Clean text data
def clean_text(text):
    # Remove "&nbsp;" and "<U+00A0>" from the beginning of the string
    text = re.sub(r'^(&nbsp;|<U\+00A0>|nbsp)+', '', text)
    # Remove special characters
    text = re.sub(r'[^\w\s]', '', text)
    # Remove non-ASCII characters
    text = re.sub(r'[^\x00-\x7F]+', '', text)
    # Convert all characters to lowercase
    text = text.lower()
    # Remove extra whitespaces
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# Apply text cleaning function to the 'text' column
cols_to_clean = ['project_description', 'sector', 'country_description', 'text.summary']
for col in cols_to_clean:
    data[col] = data[col].apply(clean_text)

# Define a function to combine columns into a new column
data['prompt'] =  data['project_description'] + ' in ' + data['sector'] + ', country ' + data['country_description'] + '\n\n###\n\n'
data['completion'] = ' ' + data['text.summary'] + 'END'

data = data[['prompt','completion']]

# Split data into training and testing sets
train_data = data.sample(frac=0.8, random_state=42)
test_data = data.drop(train_data.index)

train_data.to_json("fine-tune/train_data.jsonl", orient='records', lines=True)
test_data.to_json("fine-tune/test_data.jsonl", orient='records', lines=True)

# Format the data using fine-tuning tool
!openai tools fine_tunes.prepare_data -f fine-tune/train_data.jsonl -q
!openai tools fine_tunes.prepare_data -f fine-tune/test_data.jsonl -q

# Fine tune the model
!openai api fine_tunes.create -t "fine-tune/train_data.jsonl" -v "fine-tune/test_data_prepared.jsonl" -m davinci


