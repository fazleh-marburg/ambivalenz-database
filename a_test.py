import spacy
nlp = spacy.load("de_core_news_sm")
doc = nlp("Berlin ist die Hauptstadt von Deutschland.")
print([ent.text for ent in doc.ents])
