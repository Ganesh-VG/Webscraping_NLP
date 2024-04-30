# Web Scraping and Natural Language Processing Tailored to Client Needs
In this project, I conducted web scraping on the websites specified by the client, extracting data from each provided URL. Subsequently, I employed natural language processing techniques to analyze the extracted data according to the specific requirements outlined by the client. A screenshot showcasing the output data is provided below for reference.
![image](https://github.com/Ganesh-VG/Webscraping_NLP/assets/144704167/50852d1e-c55a-4c6e-8b47-e4c4f159b34b)
Web scraping involves fetching the HTML files of each website using the request library, followed by parsing the data using the BeautifulSoup library. The extracted data is then stored in text files for further processing, ensuring the completeness of the data extraction process.

Natural language processing (NLP) encompasses several steps. Initially, redundant punctuations are removed, followed by the elimination of stop words using a list provided by the client, stored as a text file. Subsequently, the filtered data is utilized to tally negative and positive words, leveraging lists provided by the client. From these counts, a polarity score is computed. Additionally, subjectivity scores, average sentence length, percentage of complex words, average number of words per sentence, complex word count, total word count, fog index, syllables per word, and personal pronoun usage are calculated according to the client's specified conditions and formulas.
