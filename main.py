from nltk.tokenize import sent_tokenize, word_tokenize
import pandas as pd
from bs4 import BeautifulSoup
import requests
import string
import openpyxl


def stopwords_list():
    folder_header = ["Auditor", "Currencies", "DatesandNumbers", "Generic", "GenericLong", "Geographic", "Names"]
    stopword = []

    for header in folder_header:
        with open(f"StopWords/StopWords_{header}.txt") as file:
            dummy = file.read().split("\n")
            for d in dummy:
                if "|" in d:
                    x = d.split("|")
                    stopword.append(x[0].lower().strip())
                else:
                    stopword.append(d.lower())
    return stopword

# # Activate this part of the code if you want to give the new input.
# # Remember to change the file name in pd.read_excel('<NEW FILE NAME>').
# # Remember to deactivate below function.
#
# def links():
#     dataframe1 = pd.read_excel('Input.xlsx')
#     all_keys = []
#     all_links = []
#     for per_link in range(len(dataframe1["URL"])):
#         all_links.append(dataframe1["URL"][per_link])
#         all_keys.append(dataframe1["URL_ID"][per_link])
#     return [all_links, all_keys]


# Please deactivate the function below if you have the new file to input
def to_text(data_text, keys):
    with open(f"{keys[link_no]}.txt", mode="w", encoding="utf-8") as in_data:
        in_data.write(data_text)


def links():
    all_links = ['https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-3-2/', 'https://insights.blackcoffer.com/rise-of-e-health-and-its-impact-on-humans-by-the-year-2030/', 'https://insights.blackcoffer.com/rise-of-e-health-and-its-imapct-on-humans-by-the-year-2030-2/', 'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-2/', 'https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-2-2/', 'https://insights.blackcoffer.com/rise-of-chatbots-and-its-impact-on-customer-support-by-the-year-2040/', 'https://insights.blackcoffer.com/rise-of-e-health-and-its-imapct-on-humans-by-the-year-2030/', 'https://insights.blackcoffer.com/how-does-marketing-influence-businesses-and-consumers/', 'https://insights.blackcoffer.com/how-advertisement-increase-your-market-value/', 'https://insights.blackcoffer.com/negative-effects-of-marketing-on-society/', 'https://insights.blackcoffer.com/how-advertisement-marketing-affects-business/', 'https://insights.blackcoffer.com/rising-it-cities-will-impact-the-economy-environment-infrastructure-and-city-life-by-the-year-2035/', 'https://insights.blackcoffer.com/rise-of-ott-platform-and-its-impact-on-entertainment-industry-by-the-year-2030/', 'https://insights.blackcoffer.com/rise-of-electric-vehicles-and-its-impact-on-livelihood-by-2040/', 'https://insights.blackcoffer.com/rise-of-electric-vehicle-and-its-impact-on-livelihood-by-the-year-2040/', 'https://insights.blackcoffer.com/oil-prices-by-the-year-2040-and-how-it-will-impact-the-world-economy/', 'https://insights.blackcoffer.com/an-outlook-of-healthcare-by-the-year-2040-and-how-it-will-impact-human-lives/', 'https://insights.blackcoffer.com/ai-in-healthcare-to-improve-patient-outcomes/', 'https://insights.blackcoffer.com/what-if-the-creation-is-taking-over-the-creator/', 'https://insights.blackcoffer.com/what-jobs-will-robots-take-from-humans-in-the-future/', 'https://insights.blackcoffer.com/will-machine-replace-the-human-in-the-future-of-work/', 'https://insights.blackcoffer.com/will-ai-replace-us-or-work-with-us/', 'https://insights.blackcoffer.com/man-and-machines-together-machines-are-more-diligent-than-humans-blackcoffe/', 'https://insights.blackcoffer.com/in-future-or-in-upcoming-years-humans-and-machines-are-going-to-work-together-in-every-field-of-work/', 'https://insights.blackcoffer.com/how-neural-networks-can-be-applied-in-various-areas-in-the-future/', 'https://insights.blackcoffer.com/how-machine-learning-will-affect-your-business/', 'https://insights.blackcoffer.com/deep-learning-impact-on-areas-of-e-learning/', 'https://insights.blackcoffer.com/how-to-protect-future-data-and-its-privacy-blackcoffer/', 'https://insights.blackcoffer.com/how-machines-ai-automations-and-robo-human-are-effective-in-finance-and-banking/', 'https://insights.blackcoffer.com/ai-human-robotics-machine-future-planet-blackcoffer-thinking-jobs-workplace/', 'https://insights.blackcoffer.com/how-ai-will-change-the-world-blackcoffer/', 'https://insights.blackcoffer.com/future-of-work-how-ai-has-entered-the-workplace/', 'https://insights.blackcoffer.com/ai-tool-alexa-google-assistant-finance-banking-tool-future/', 'https://insights.blackcoffer.com/ai-healthcare-revolution-ml-technology-algorithm-google-analytics-industrialrevolution/', 'https://insights.blackcoffer.com/all-you-need-to-know-about-online-marketing/', 'https://insights.blackcoffer.com/evolution-of-advertising-industry/', 'https://insights.blackcoffer.com/how-data-analytics-can-help-your-business-respond-to-the-impact-of-covid-19/', 'https://insights.blackcoffer.com/covid-19-environmental-impact-for-the-future/', 'https://insights.blackcoffer.com/environmental-impact-of-the-covid-19-pandemic-lesson-for-the-future/', 'https://insights.blackcoffer.com/how-data-analytics-and-ai-are-used-to-halt-the-covid-19-pandemic/', 'https://insights.blackcoffer.com/difference-between-artificial-intelligence-machine-learning-statistics-and-data-mining/', 'https://insights.blackcoffer.com/how-python-became-the-first-choice-for-data-science/', 'https://insights.blackcoffer.com/how-google-fit-measure-heart-and-respiratory-rates-using-a-phone/', 'https://insights.blackcoffer.com/what-is-the-future-of-mobile-apps/', 'https://insights.blackcoffer.com/impact-of-ai-in-health-and-medicine/', 'https://insights.blackcoffer.com/telemedicine-what-patients-like-and-dislike-about-it/', 'https://insights.blackcoffer.com/how-we-forecast-future-technologies/', 'https://insights.blackcoffer.com/can-robots-tackle-late-life-loneliness/', 'https://insights.blackcoffer.com/embedding-care-robots-into-society-socio-technical-considerations/', 'https://insights.blackcoffer.com/management-challenges-for-future-digitalization-of-healthcare-services/', 'https://insights.blackcoffer.com/are-we-any-closer-to-preventing-a-nuclear-holocaust/', 'https://insights.blackcoffer.com/will-technology-eliminate-the-need-for-animal-testing-in-drug-development/', 'https://insights.blackcoffer.com/will-we-ever-understand-the-nature-of-consciousness/', 'https://insights.blackcoffer.com/will-we-ever-colonize-outer-space/', 'https://insights.blackcoffer.com/what-is-the-chance-homo-sapiens-will-survive-for-the-next-500-years/', 'https://insights.blackcoffer.com/why-does-your-business-need-a-chatbot/', 'https://insights.blackcoffer.com/how-you-lead-a-project-or-a-team-without-any-technical-expertise/', 'https://insights.blackcoffer.com/can-you-be-great-leader-without-technical-expertise/', 'https://insights.blackcoffer.com/how-does-artificial-intelligence-affect-the-environment/', 'https://insights.blackcoffer.com/how-to-overcome-your-fear-of-making-mistakes-2/', 'https://insights.blackcoffer.com/is-perfection-the-greatest-enemy-of-productivity/', 'https://insights.blackcoffer.com/global-financial-crisis-2008-causes-effects-and-its-solution/', 'https://insights.blackcoffer.com/gender-diversity-and-equality-in-the-tech-industry/', 'https://insights.blackcoffer.com/how-to-overcome-your-fear-of-making-mistakes/', 'https://insights.blackcoffer.com/how-small-business-can-survive-the-coronavirus-crisis/', 'https://insights.blackcoffer.com/impacts-of-covid-19-on-vegetable-vendors-and-food-stalls/', 'https://insights.blackcoffer.com/impacts-of-covid-19-on-vegetable-vendors/', 'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-tourism-aviation-industries/', 'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-sports-events-around-the-world/', 'https://insights.blackcoffer.com/changing-landscape-and-emerging-trends-in-the-indian-it-ites-industry/', 'https://insights.blackcoffer.com/online-gaming-adolescent-online-gaming-effects-demotivated-depression-musculoskeletal-and-psychosomatic-symptoms/', 'https://insights.blackcoffer.com/human-rights-outlook/', 'https://insights.blackcoffer.com/how-voice-search-makes-your-business-a-successful-business/', 'https://insights.blackcoffer.com/how-the-covid-19-crisis-is-redefining-jobs-and-services/', 'https://insights.blackcoffer.com/how-to-increase-social-media-engagement-for-marketers/', 'https://insights.blackcoffer.com/impacts-of-covid-19-on-streets-sides-food-stalls/', 'https://insights.blackcoffer.com/coronavirus-impact-on-energy-markets-2/', 'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry-5/', 'https://insights.blackcoffer.com/lessons-from-the-past-some-key-learnings-relevant-to-the-coronavirus-crisis-4/', 'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work-2/', 'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work-3/', 'https://insights.blackcoffer.com/travel-and-tourism-outlook/', 'https://insights.blackcoffer.com/gaming-disorder-and-effects-of-gaming-on-health/', 'https://insights.blackcoffer.com/what-is-the-repercussion-of-the-environment-due-to-the-covid-19-pandemic-situation/', 'https://insights.blackcoffer.com/what-is-the-repercussion-of-the-environment-due-to-the-covid-19-pandemic-situation-2/', 'https://insights.blackcoffer.com/impact-of-covid-19-pandemic-on-office-space-and-co-working-industries/', 'https://insights.blackcoffer.com/contribution-of-handicrafts-visual-arts-literature-in-the-indian-economy/', 'https://insights.blackcoffer.com/how-covid-19-is-impacting-payment-preferences/', 'https://insights.blackcoffer.com/how-will-covid-19-affect-the-world-of-work-2/', 'https://insights.blackcoffer.com/lessons-from-the-past-some-key-learnings-relevant-to-the-coronavirus-crisis/', 'https://insights.blackcoffer.com/covid-19-how-have-countries-been-responding/', 'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry-2/', 'https://insights.blackcoffer.com/how-will-covid-19-affect-the-world-of-work-3/', 'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry-3/', 'https://insights.blackcoffer.com/estimating-the-impact-of-covid-19-on-the-world-of-work/', 'https://insights.blackcoffer.com/covid-19-how-have-countries-been-responding-2/', 'https://insights.blackcoffer.com/how-will-covid-19-affect-the-world-of-work-4/', 'https://insights.blackcoffer.com/lessons-from-the-past-some-key-learnings-relevant-to-the-coronavirus-crisis-2/', 'https://insights.blackcoffer.com/lessons-from-the-past-some-key-learnings-relevant-to-the-coronavirus-crisis-3/', 'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry-4/', 'https://insights.blackcoffer.com/why-scams-like-nirav-modi-happen-with-indian-banks/', 'https://insights.blackcoffer.com/impact-of-covid-19-on-the-global-economy/', 'https://insights.blackcoffer.com/impact-of-covid-19coronavirus-on-the-indian-economy-2/', 'https://insights.blackcoffer.com/impact-of-covid-19-on-the-global-economy-2/', 'https://insights.blackcoffer.com/impact-of-covid-19-coronavirus-on-the-indian-economy-3/', 'https://insights.blackcoffer.com/should-celebrities-be-allowed-to-join-politics/', 'https://insights.blackcoffer.com/how-prepared-is-india-to-tackle-a-possible-covid-19-outbreak/', 'https://insights.blackcoffer.com/how-will-covid-19-affect-the-world-of-work/', 'https://insights.blackcoffer.com/controversy-as-a-marketing-strategy/', 'https://insights.blackcoffer.com/coronavirus-impact-on-the-hospitality-industry/', 'https://insights.blackcoffer.com/coronavirus-impact-on-energy-markets/', 'https://insights.blackcoffer.com/what-are-the-key-policies-that-will-mitigate-the-impacts-of-covid-19-on-the-world-of-work/', 'https://insights.blackcoffer.com/marketing-drives-results-with-a-focus-on-problems/', 'https://insights.blackcoffer.com/continued-demand-for-sustainability/']
    all_keys = ['123.0', '321.0', '2345.0', '4321.0', '432.0', '2893.8', '3355.6', '3817.4', '4279.2', '4741.0', '5202.8', '5664.6', '6126.4', '6588.2', '7050.0', '7511.8', '7973.6', '8435.4', '8897.2', '9359.0', '9820.8', '10282.6', '10744.4', '11206.2', '11668.0', '12129.8', '12591.6', '13053.4', '13515.2', '13977.0', '14438.8', '14900.6', '15362.4', '15824.2', '16286.0', '16747.8', '17209.6', '17671.4', '18133.2', '18595.0', '19056.8', '19518.6', '19980.4', '20442.2', '20904.0', '21365.8', '21827.6', '22289.4', '22751.2', '23213.0', '23674.8', '24136.6', '24598.4', '25060.2', '25522.0', '25983.8', '26445.6', '26907.4', '27369.2', '27831.0', '28292.8', '28754.6', '29216.4', '29678.2', '30140.0', '30601.8', '31063.6', '31525.4', '31987.2', '32449.0', '32910.8', '33372.6', '33834.4', '34296.2', '34758.0', '35219.8', '35681.6', '36143.4', '36605.2', '37067.0', '37528.8', '37990.6', '38452.4', '38914.2', '39376.0', '39837.8', '40299.6', '40761.4', '41223.2', '41685.0', '42146.8', '42608.6', '43070.4', '43532.2', '43994.0', '44455.8', '44917.6', '45379.4', '45841.2', '46303.0', '46764.8', '47226.6', '47688.4', '48150.2', '48612.0', '49073.8', '49535.6', '49997.4', '50459.2', '50921.0', '51382.8', '51844.6', '52306.4', '52768.2']
    return [all_links, all_keys]


def web_data(link_web):
    response = requests.get(f"{link_web}")
    web_page = response.text
    soup = BeautifulSoup(web_page, "html.parser")
    title = soup.find(name="h1")
    text = title.text
    parent = soup.find(name="body").find_all(name="p")
    new_parent = parent[16:-3:1]
    for all_text in new_parent:
        text += " "
        text += all_text.getText()
    return text


def filter_statement(data_in):
    punct = list(string.punctuation)
    for ele in data_in:
        if ele in punct:
            data_in = data_in.replace(ele, "")
    stop_words = stopwords_list()
    word_tokens = word_tokenize(data_in)
    filtered_sentence = [w for w in word_tokens if not w.lower() in stop_words]
    return filtered_sentence


def positive_word_fun(positive_statement):
    with open("MasterDictionary-20231010T132545Z-001/MasterDictionary/positive-words.txt") as positive_data:
        p_data = positive_data.read().split("\n")
    no_positive_words = 0
    for find_p_word in positive_statement:
        if find_p_word in p_data:
            no_positive_words += 1
    return no_positive_words


def negative_word_fun(negative_statement):
    with open("MasterDictionary-20231010T132545Z-001/MasterDictionary/negative-words.txt") as negative_data:
        n_data = negative_data.read().split("\n")
    no_negative_words = 0
    for find_p_word in negative_statement:
        if find_p_word in n_data:
            no_negative_words += 1
    return no_negative_words


def complex_word_counter(complex_word_data):
    vowels = list("aeiou")
    complex_word_count = 0
    complex_word_data_list = complex_word_data.split()
    for each_word in complex_word_data_list:
        count_vowels = 0
        each_word_list = list(each_word.lower())
        for each_alphabet in each_word_list:
            if each_alphabet in vowels:
                count_vowels += 1
                if ((each_word_list[-1] == "d" and each_word_list[-2] == "e") or
                        (each_word_list[-1] == "s" and each_word_list[-2] == "e")):
                    count_vowels -= 1
                if count_vowels > 2:
                    complex_word_count += 1
    return complex_word_count


def syllable_counter(syllable_data):
    vowels = list("aeiou")
    syllable_count = 0
    syllable_data_list = syllable_data.split()
    for each_word in syllable_data_list:
        each_word_list = list(each_word.lower())
        for each_alphabet in each_word_list:
            if each_alphabet in vowels:
                syllable_count += 1
                if ((each_word_list[-1] == "d" and each_word_list[-2] == "e") or
                        (each_word_list[-1] == "s" and each_word_list[-2] == "e")):
                    syllable_count -= 1
    return syllable_count


def personal_pronoun_counter(personal_pronoun_data):
    personal_pronoun_count = 0
    personal_pronoun_list = ["i", "we", "my", "ours"]
    personal_pronoun_data_list = personal_pronoun_data.split()
    for each_word_pronoun in personal_pronoun_data_list:
        if each_word_pronoun.lower() in personal_pronoun_list or (each_word_pronoun.lower() == "us" and each_word_pronoun.isupper() == False):
            personal_pronoun_count += 1
    return personal_pronoun_count


def total_characters_calculator(total_characters_data):
    punct = list(string.punctuation)
    total_characters_count = 0
    total_characters_data_list = total_characters_data.split()
    for each_word_total_characters in total_characters_data_list:
        each_word_total_characters_list = list(each_word_total_characters)
        for character in each_word_total_characters_list:
            if character not in punct:
                total_characters_count += 1
    return total_characters_count


def word_count(raw_data_wor):
    punct = list(string.punctuation)
    for ele in raw_data_wor:
        if ele in punct:
            raw_data_wor = raw_data_wor.replace(ele, "")
    return raw_data_wor


link_keys = links()
links_out = link_keys[0]
keys_out = link_keys[1]
dataframe_input = []
dataframe_index = []
dataframe_columns = ["URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
                     "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "AVG NUMBER OF WORDS PER SENTENCE",
                     "COMPLEX WORD COUNT", "WORD COUNT", "FOG INDEX", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"]

for link_no in range(len(links_out)):
    input_unit = []
    dataframe_index.append(link_no)
    input_unit.append(keys_out[link_no])
    input_unit.append(links_out[link_no])
    print(links_out[link_no])
    try:
        web_data_out = web_data(links_out[link_no])
    except AttributeError:
        for _ in range(0, 13):
            input_unit.append("Nil")
        dataframe_input.append(input_unit)
        continue

    # Save data to text files
    to_text(web_data_out, keys_out)

    # Filtered statement after the removal of stop words
    filter_statement_out = filter_statement(web_data_out)
    # print(f"filter_statement_out: {filter_statement_out}")

    # Positive Score
    positive_words = positive_word_fun(filter_statement_out)
    input_unit.append(positive_words)
    print(f"positive_words: {positive_words}")

    # Negative Score
    negative_words = negative_word_fun(filter_statement_out)
    input_unit.append(negative_words)
    print(f"negative_words: {negative_words}")

    # Number of words after cleaning
    no_of_words_cleaned = len(filter_statement_out)
    print(f"no_of_words_cleaned: {no_of_words_cleaned}")

    # Polarity Score = (Positive Score – Negative Score)/ ((Positive Score + Negative Score) + 0.000001)
    polarity_score = (positive_words - negative_words) / (positive_words + negative_words + 0.000001)
    input_unit.append(polarity_score)
    print(f"polarity_score: {polarity_score}")

    # Subjectivity Score = (Positive Score + Negative Score)/ ((Total Words after cleaning) + 0.000001)
    subjective_score = (positive_words + negative_words) / (no_of_words_cleaned + 0.000001)
    input_unit.append(subjective_score)
    print(f"subjective_score: {subjective_score}")

    # Raw web data after all the punctuations are removed
    punct_cleared_data = word_count(web_data_out)

    # number of words in raw data extracted from web
    no_of_words = len(punct_cleared_data.split())
    print(f"no_of_words: {no_of_words}")

    # number of sentences in raw data extracted from web
    sentence_tokens = sent_tokenize(web_data_out)
    no_of_sentences = len(sentence_tokens)
    print(f"no_of_sentences: {no_of_sentences}")

    # Average Sentence Length = the number of words / the number of sentences
    average_sentence_length = no_of_words / no_of_sentences
    input_unit.append(average_sentence_length)
    print(f"average_sentence_length: {average_sentence_length}")

    # the number of complex words in raw data extracted from web
    no_of_complex_words = complex_word_counter(web_data_out)
    print(f"no_of_complex_words: {no_of_complex_words}")

    # Percentage of Complex words = the number of complex words / the number of words
    percentage_of_complex_words = no_of_complex_words / no_of_words
    input_unit.append(percentage_of_complex_words)
    print(f"percentage_of_complex_words: {percentage_of_complex_words}")

    # Average Number of Words Per Sentence = the total number of words / the total number of sentences
    average_number_words_per_sentence = no_of_words / no_of_sentences
    input_unit.append(average_number_words_per_sentence)
    print(f"average_number_words_per_sentence: {average_number_words_per_sentence}")

    input_unit.append(no_of_complex_words)
    input_unit.append(no_of_words)

    # Fog Index = 0.4 * (Average Sentence Length + Percentage of Complex words)
    fog_index = 0.4 * (average_sentence_length + percentage_of_complex_words)
    input_unit.append(fog_index)
    print(f"fog_index: {fog_index}")

    #  number of Syllables in each word of the text by counting the vowels present in each word
    no_syllable_count = syllable_counter(punct_cleared_data)
    input_unit.append(no_syllable_count)
    print(f"no_syllable_count: {no_syllable_count}")

    # To calculate Personal Pronouns mentioned in the text
    no_personal_pronoun = personal_pronoun_counter(punct_cleared_data)
    input_unit.append(no_personal_pronoun)
    print(f"no_personal_pronoun: {no_personal_pronoun}")

    # Sum of the total number of characters in each word
    total_no_of_characters = total_characters_calculator(punct_cleared_data)
    print(f"total_no_of_characters: {total_no_of_characters}")

    # Average Word Length is calculated by the formula:
    # Sum of the total number of characters in each word/Total number of words
    average_word_length = total_no_of_characters / no_of_words
    input_unit.append(average_word_length)
    print(f"average_word_length: {average_word_length}")

    dataframe_input.append(input_unit)

print(dataframe_input)
for index_no in range(len(links_out)):
    excel_df = pd.DataFrame(dataframe_input, index=dataframe_index, columns=dataframe_columns)

excel_df.to_excel( "Output.xlsx", sheet_name="sheet1")