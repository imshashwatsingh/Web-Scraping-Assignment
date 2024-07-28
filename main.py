import pandas as pd
import string
import re
# import os

# Cleaning !
f1 = open("StopWords\StopWords_Auditor.txt")
filter1 = f1.read().split()
# print(filter1)

filter2 = []

try:
    with open("StopWords\StopWords_Currencies.txt", 'r') as f2:
        while True:
            line = f2.readline().split()
            if not line:
                break
            for x in line:
                if x != '|':
                    filter2.append(x)
            # print(line, end='')  # end='' to avoid adding extra new lines
except FileNotFoundError:
    print(f"The file was not found.")
except IOError:
    print(f"An error occurred while reading the file.")

# print(filter2)

filter3 = []

try:
    with open("StopWords\StopWords_DatesandNumbers.txt", 'r') as f3:
        while True:
            line = f3.readline().split()
            if not line:
                break
            for x in line:
                if x != '|':
                    filter3.append(x)
            # print(line, end='')  # end='' to avoid adding extra new lines
except FileNotFoundError:
    print(f"The file was not found.")
except IOError:
    print(f"An error occurred while reading the file.")

# print(filter3)

f4 = open("StopWords\StopWords_Generic.txt")
filter4 = f4.read().split()
# print(filter4)

f5 = open("StopWords\StopWords_GenericLong.txt")
filter5 = f5.read().split()
# print(filter5)

filter6 = []

try:
    with open("StopWords\StopWords_Geographic.txt", 'r') as f6:
        while True:
            line = f6.readline().lower().split()
            if not line:
                break
            for x in line:
                if x != '|':
                    filter6.append(x)
            # print(line, end='')  # end='' to avoid adding extra new lines
except FileNotFoundError:
    print(f"The file was not found.")
except IOError:
    print(f"An error occurred while reading the file.")

# print(filter6)

filter7 = []

f7 = open("StopWords\StopWords_Names.txt")

lines = f7.readlines()

lines = lines

for line in lines:
    words = line.split()
    if words:
        filter7.append(words[0].lower())

# print(filter7)

clean = [filter1,filter2,filter3,filter4,filter5,filter6,filter7]

# Fucntions

def filter(original_list, strings_to_remove):
    return [string for string in original_list if string not in strings_to_remove]

def remove_punctuation_from_strings(string_list):
    translator = str.maketrans('', '', string.punctuation)
    return [s.translate(translator) for s in string_list]

def count_lines(text):
    lines = text.splitlines()
    line_count = 0
    if lines:
        for line in lines:
            if line.strip():
                line_count += 1
        return line_count
    return 0


def count_words(text):
    words = text.split()
    if words:
        return len(words) 
    else:
        return 0

def count_sentences(text):
    sentence_endings = re.compile(r'[.!?]')
    
    sentences = sentence_endings.split(text)
    
    sentence_count = sum(1 for sentence in sentences if sentence.strip())
    
    return sentence_count


def count_syllables(word):
    word = word.lower().strip()
    
    # Handle special cases for 'es' and 'ed'
    if word.endswith('es') and len(word) > 2 and word[-3] not in 'aeiouy':
        word = word[:-2]
    elif word.endswith('ed') and len(word) > 2 and word[-3] not in 'aeiouy':
        word = word[:-2]
    
    # Remove non-alphabetic characters for cleaner syllable counting
    word = re.sub(r'[^a-z]', '', word)
    
    # Count the number of vowels in the word
    syllable_count = len(re.findall(r'[aeiouy]', word))
    
    # Ensure at least one syllable for every word
    if syllable_count == 0:
        syllable_count = 1
    
    return syllable_count

def syllables_in_text(text):
    count = 0
    words = re.findall(r'\b\w+\b', text)
    for word in words:
        count+=  count_syllables(word)
    return count


def count_complex_words(string_list):
    count = 0
    # complex_words = []
    for sentence in string_list:
        words = sentence.split()
        for word in words:
            syllable_count = count_syllables(word)
            if syllable_count > 2:
                # complex_words.append(word)
                count+=1
    return count

def count_personal_pronouns(text):
    # Define a regex pattern for personal pronouns
    pronoun_pattern = re.compile(r'\b(I|we|my|ours|us)\b', re.IGNORECASE)
    
    # Find all matches in the text
    matches = pronoun_pattern.findall(text)
    
    # Filter out 'US' that should not be counted
    filtered_matches = [match for match in matches if match.lower() != 'us' or re.search(r'\bUS\b', text) is None]
    
    return len(filtered_matches)

def average_word_length(text):
    words = re.findall(r'\b\w+\b', text)
    
    total_words = len(words)
    
    total_characters = sum(len(word) for word in words)
   
    if total_words == 0:
        return 0  
    average_length = total_characters / total_words

    return average_length


output = pd.read_excel("IOFiles//Output Data Structure.xlsx")

p = open("MasterDictionary/positive-words.txt","r")
positive_words = p.read().lower().split()

p = open("MasterDictionary/negative-words.txt","r")
negative_words = p.read().lower().split()


def per_comp_words(text):
    comp_words = count_complex_words(text)
    num_of_words = count_words(text)
    if num_of_words !=0:
        return comp_words/num_of_words
    return 0

def avg_sen_len(text):
    num_words = count_words(text)
    num_sen = count_sentences(text)
    if num_sen!=0:
        return num_words/num_sen
    return 0


for i in range(output.shape[0]):
    try:
        f = open("articles/"+str(output.iloc[i,0])+".txt","r",encoding='utf-8')
        text = f.read().lower().strip()
        f.seek(0)
        text2 = f.read().strip()

        for x in range(7):
            data = filter(text,clean[x])


        data = remove_punctuation_from_strings(text)

        total_words_after_cleaning = len(data)

        # positive_score
        positive_score = 0
        for xyz in data:
            if xyz in positive_words:
                positive_score +=1

        # negative_score
        negative_score = 0
        for xn in data:
            if xn in negative_words:
                negative_score -=1

        negative_score *= -1

        polarity_score =(positive_score - negative_score)/((positive_score + negative_score) + 0.000001)

        subjectivity_score = (positive_score+negative_score)/((total_words_after_cleaning)+0.000001)

        number_of_lines = count_lines(text)
        number_of_words = count_words(text)
        number_of_sentences = count_sentences(text)
        complex_word_count = count_complex_words(text)
        number_of_syllables_per_word = syllables_in_text(text)
        personal_prounouns = count_personal_pronouns(text)
        avg_word_length = average_word_length(text)
        avg_sentence_length = avg_sen_len(text)
        per_conmplex_words = per_comp_words(text)
        fog_index = 0.4*(avg_sentence_length+per_conmplex_words)
        avg_number_of_words_per_sentence = avg_sen_len(text)
        row = [positive_score,negative_score,polarity_score,subjectivity_score,avg_sentence_length,per_conmplex_words,fog_index,avg_number_of_words_per_sentence,complex_word_count,total_words_after_cleaning,number_of_syllables_per_word,personal_prounouns,avg_word_length]
        print(i+1,end=', ')
        for j in range(2,len(row)+2):
            output.iloc[i,j] = row[j-2]
    except Exception as e:
        continue

output.to_excel("output.xlsx")
print("Output File created successfully !")