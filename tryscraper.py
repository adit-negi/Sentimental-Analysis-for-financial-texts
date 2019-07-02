import requests
import urllib
from bs4 import BeautifulSoup
import time
import sys
# Reading an excel file using Python
import xlrd
import xlsxwriter
import xlwt
from xlwt import Workbook

# Give the location of the file
loc = ('/home/adit/cik_list.xlsx')
locw = ('/home/adit/OutputDataStructure.xlsx')
# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
workbook = xlsxwriter.Workbook('Example.xlsx')
sheet1 = workbook.add_worksheet()
sys.setrecursionlimit(100000)
# For row 0 and column 0
for i in range(1, 152):
    a = sheet.cell_value(i, 5)
    url = ('https://www.sec.gov/Archives/'+a)

    response = requests.get(url)
    #scraping the pages one by one

    soup = BeautifulSoup(response.text, "html.parser")
    txt = str(soup)
    text = txt.lower()

    doc_lenght = len(text)
    #function to extract mda text

    def mdafc(text, doc_lenght):
        first_instance_part1 = text.find(
            "management's discussion and")+1

        if first_instance_part1 != 0:    # index page reference

            new_text = text[(first_instance_part1)*3:doc_lenght]                 #skipping the idex
            second_instance_part1 = new_text.find(
                "management's discussion and")                                  #title reference

            text_forFlag = text[(first_instance_part1)*3 +
                                second_instance_part1:doc_lenght]
            flag = text_forFlag.find("item ")                                    

            try:
                mda1 = (text_forFlag[flag+5])                                     # to check if its a title not a word
                mda = (text_forFlag[:flag])

                return mda
            except IndexError:
                return 0
        else:
            return 0
    #function to extract qqdmr text
    def qqdmr(text, doc_lenght):
        first_instance_part2 = text.find(
            "quantitative and qualitative disclosures")+1

        if first_instance_part2 != 0:

            new_text = text[(first_instance_part2)*3:doc_lenght]
            second_instance_part2 = new_text.find(
                "management's discussion and")

            text_forFlag = text[(first_instance_part2)*3 +
                                second_instance_part2:doc_lenght]
            flag = text_forFlag.find("item ")
            try:
                qqd1 = (text_forFlag[flag+5])
                qqd = (text_forFlag[:flag])

                return qqd
            except IndexError:
                return 0
        else:
            return 0

    def rf(text, doc_lenght):

        first_instance_part3 = text.find("risk factors")+1

        if first_instance_part3 != 0:

            new_text = text[(first_instance_part3)*3:doc_lenght]
            second_instance_part3 = new_text.find(
                "management's discussion and")

            text_forFlag = text[(first_instance_part3)*3 +
                                second_instance_part3:doc_lenght]
            flag = text_forFlag.find("item ")

            rfs = (text_forFlag[:flag])
            return rfs

        else:
            return 0

    def cleanText(Ctext):
        """
        removes punctuation, stopwords and returns lowercase text in a list of single words
        """
        Ctext = Ctext.lower()

        from bs4 import BeautifulSoup
        Ctext = BeautifulSoup(Ctext, features="lxml").get_text()

        from nltk.tokenize import RegexpTokenizer
        tokenizer = RegexpTokenizer(r'\w+')
        Ctext = tokenizer.tokenize(Ctext)

        from nltk.corpus import stopwords
        clean = [
            word for word in Ctext if word not in stopwords.words('english')]

        return clean

    def loadPositive():
        """
        loading positive dictionary
        """
        myfile = open('/home/adit/positive.csv', "r")
        positives = myfile.readlines()
        positive = [pos.strip().lower() for pos in positives]
        return positive

    def loadNegative():
        """
        loading positive dictionary
        """
        myfile = open('/home/adit/negative.csv', "r")
        negatives = myfile.readlines()
        negative = [neg.strip().lower() for neg in negatives]
        return negative

    def loadConstrain():
        """
        loading constraining dictionary
        """
        myfile = open('/home/adit/constrain.csv', "r")
        constrains = myfile.readlines()
        constrain = [con.strip().lower() for con in constrains]
        return constrain

    def loadUncertain():
        """
        loading uncertainity dictionary
        """
        myfile = open('/home/adit/uncertain.csv', "r")
        uncertains = myfile.readlines()
        uncertain = [un.strip().lower() for un in uncertains]
        return uncertain

    def countNeg(cleantext, negative):
        """
        counts negative words in cleantext
        """
        negs = [word for word in cleantext if word in negative]
        return len(negs)

    def countPos(cleantext, positive):
        """
        counts negative words in cleantext
        """
        pos = [word for word in cleantext if word in positive]
        return len(pos)

    def countCons(cleantext, constrain):
        """
        counts negative words in cleantext
        """
        con = [word for word in cleantext if word in constrain]
        return len(con)

    def countUn(cleantext, uncertain):
        """
        counts negative words in cleantext
        """
        un = [word for word in cleantext if word in uncertain]
        return len(un)

    def getSentiment(cleantext, negative, positive):
        """
        counts negative and positive words in cleantext and returns a score accordingly
        """
        positive = loadPositive()
        negative = loadNegative()
        return ((countPos(cleantext, positive) - countNeg(cleantext, negative))/(countPos(cleantext, positive) + countNeg(cleantext, negative) + 0.000001))

    def complex_word(words):
        ctr = 0

        for word in words:
            if len(word) > 2:
                ctr += 1
        return ctr
    #initializing
    mda = mdafc(text, doc_lenght)
    qqd = qqdmr(text, doc_lenght)
    rfs = rf(text, doc_lenght)
    positive = loadPositive()
    negative = loadNegative()
    constrain = loadConstrain()
    uncertain = loadUncertain()
    total_cons = countCons(text, constrain)

    #considering all the cases(i made a mistake while reading the docs)
    if mda != 0 and qqd == 0 and rfs == 0:

        print(1)
        obj_clean = cleanText(mda)
        mda_positive_count = countPos(obj_clean, positive)
        mda_negative_count = countNeg(obj_clean, negative)
        mda_constrain_count = countCons(obj_clean, constrain)
        mda_uncertain_count = countUn(obj_clean, uncertain)
        mda_polarity = getSentiment(obj_clean, negative, positive)
        mda_word_count = len(obj_clean)
        mda_avg_sentence_lenght = len(mda.split())/mda.count('.')
        mda_complex_words = complex_word(obj_clean)
        mda_percentage_complex_words = (mda_complex_words/mda_word_count)*100
        mda_fog_idex = 0.4*(mda_avg_sentence_lenght +
                            mda_percentage_complex_words)
        mda_positive_ratio = mda_positive_count/mda_word_count
        mda_negative_ratio = mda_negative_count/mda_word_count
        mda_uncertain_ratio = mda_uncertain_count/mda_word_count
        mda_constrain_ratio = mda_constrain_count/mda_word_count
        qqd_positive_count = ""
        qqd_negative_count = ""
        qqd_constrain_count = ""
        qqd_uncertain_count = ""
        qqd_polarity = ""
        qqd_word_count = ""
        qqd_avg_sentence_lenght = ""
        qqd_complex_words = ""
        qqd_percentage_complex_words = ""
        qqd_fog_idex = ""
        qqd_positive_ratio = ""
        qqd_negative_ratio = ""
        qqd_uncertain_ratio = ""
        qqd_constrain_ratio = ""
        rf_positive_count = ""
        rf_negative_count = ""
        rf_constrain_count = ""
        rf_uncertain_count = ""
        rf_polarity = ""
        rf_word_count = ""

        rf_avg_sentence_lenght = ""
        rf_complex_words = ""
        rf_percentage_complex_words = ""
        rf_fog_idex = ""
        rf_positive_ratio = ""
        rf_negative_ratio = ""
        rf_uncertain_ratio = ""
        rf_constrain_ratio = ""

    elif mda == 0 and qqd != 0 and rfs == 0:

        print(2)
        obj_clean = cleanText(qqd)
        qqd_positive_count = countPos(obj_clean, positive)
        qqd_negative_count = countNeg(obj_clean, negative)
        qqd_constrain_count = countCons(obj_clean, constrain)
        qqd_uncertain_count = countUn(obj_clean, uncertain)
        qqd_polarity = getSentiment(obj_clean, negative, positive)
        qqd_word_count = len(obj_clean)
        qqd_avg_sentence_lenght = len(qqd.split())/qqd.count('.')
        qqd_complex_words = complex_word(obj_clean)
        qqd_percentage_complex_words = (qqd_complex_words/qqd_word_count)*100
        qqd_fog_idex = 0.4*(qqd_avg_sentence_lenght +
                            qqd_percentage_complex_words)
        qqd_positive_ratio = qqd_positive_count/qqd_word_count
        qqd_negative_ratio = qqd_negative_count/qqd_word_count
        qqd_uncertain_ratio = qqd_uncertain_count/qqd_word_count
        qqd_constrain_ratio = qqd_constrain_count/qqd_word_count
        rf_positive_count = ""
        rf_negative_count = ""
        rf_constrain_count = ""
        rf_uncertain_count = ""
        rf_polarity = ""
        rf_word_count = ""
        rf_avg_sentence_lenght = ""
        rf_complex_words = ""
        rf_percentage_complex_words = ""
        rf_fog_idex = ""
        rf_positive_ratio = ""
        rf_negative_ratio = ""
        rf_uncertain_ratio = ""
        rf_constrain_ratio = ""
        mda_positive_count = ""
        mda_negative_count = ""
        mda_constrain_count = ""
        mda_uncertain_count = ""
        mda_polarity = ""
        mda_word_count = ""
        mda_avg_sentence_lenght = ""
        mda_complex_words = ""
        mda_percentage_complex_words = ""
        mda_fog_idex = ""
        mda_positive_ratio = ""
        mda_negative_ratio = ""
        mda_uncertain_ratio = ""
        mda_constrain_ratio = ""

    elif mda == 0 and qqd == 0 and rfs != 0:

        print(3)
        obj_clean = cleanText(rfs)
        rf_positive_count = countPos(obj_clean, positive)
        rf_negative_count = countNeg(obj_clean, negative)
        rf_constrain_count = countCons(obj_clean, constrain)
        rf_uncertain_count = countUn(obj_clean, uncertain)
        rf_polarity = getSentiment(obj_clean, negative, positive)
        rf_word_count = len(obj_clean)
        if rfs.count('.') == 0:
            rf_avg_sentence_lenght = 0
        else:
            rf_avg_sentence_lenght = len(rfs.split())/rfs.count('.')
        rf_complex_words = complex_word(obj_clean)
        if rf_word_count == 0:
            rf_percentage_complex_words = ""
            rf_positive_ratio = ""
            rf_negative_ratio = ""
            rf_uncertain_ratio = ""
            rf_constrain_ratio = ""
            rf_fog_idex = ""
        else:
            rf_percentage_complex_words = (rf_complex_words/rf_word_count)*100

            rf_positive_ratio = rf_positive_count/rf_word_count
            rf_negative_ratio = rf_negative_count/rf_word_count
            rf_uncertain_ratio = rf_uncertain_count/rf_word_count
            rf_constrain_ratio = rf_constrain_count/rf_word_count
            rf_fog_idex = 0.4*(rf_avg_sentence_lenght +
                               rf_percentage_complex_words)

        mda_positive_count = ""
        mda_negative_count = ""
        mda_constrain_count = ""
        mda_uncertain_count = ""
        mda_polarity = ""
        mda_word_count = ""
        mda_avg_sentence_lenght = ""
        mda_complex_words = ""
        mda_percentage_complex_words = ""
        mda_fog_idex = ""
        mda_positive_ratio = ""
        mda_negative_ratio = ""
        mda_uncertain_ratio = ""
        mda_constrain_ratio = ""

        qqd_positive_count = ""
        qqd_negative_count = ""
        qqd_constrain_count = ""
        qqd_uncertain_count = ""
        qqd_polarity = ""
        qqd_word_count = ""
        qqd_avg_sentence_lenght = ""
        qqd_complex_words = ""
        qqd_percentage_complex_words = ""
        qqd_fog_idex = ""
        qqd_positive_ratio = ""
        qqd_negative_ratio = ""
        qqd_uncertain_ratio = ""
        qqd_constrain_ratio = ""

    elif mda != 0 and qqd != 0 and rfs == 0:

        print(4)
        obj_clean_mda = cleanText(mda)
        obj_clean_qqd = cleanText(qqd)

        mda_positive_count = countPos(obj_clean_mda, positive)
        qqd_positive_count = countPos(obj_clean_qqd, positive)
        qqd_negative_count = countNeg(obj_clean_qqd, negative)
        mda_constrain_count = countCons(obj_clean_mda, constrain)
        qqd_constrain_count = countCons(obj_clean_qqd, constrain)
        mda_uncertain_count = countUn(obj_clean_mda, uncertain)
        qqd_uncertain_count = countUn(obj_clean_qqd, uncertain)
        mda_polarity = getSentiment(obj_clean_mda, negative, positive)
        qqd_polarity = getSentiment(obj_clean_qqd, negative, positive)
        mda_word_count = len(obj_clean_mda)
        qqd_word_count = len(obj_clean_qqd)
        mda_avg_sentence_lenght = len(qqd.split())/mda.count('.')
        qqd_avg_sentence_lenght = len(mda.split())/qqd.count('.')
        mda_complex_words = complex_word(obj_clean_mda)
        qqd_complex_words = complex_word(obj_clean_qqd)
        mda_percentage_complex_words = (mda_complex_words/mda_word_count)*100
        qqd_percentage_complex_words = (qqd_complex_words/qqd_word_count)*100
        mda_fog_idex = 0.4*(mda_avg_sentence_lenght +
                            mda_percentage_complex_words)
        qqd_fog_idex = 0.4*(qqd_avg_sentence_lenght +
                            qqd_percentage_complex_words)

        qqd_positive_ratio = qqd_positive_count/qqd_word_count
        qqd_negative_ratio = qqd_negative_count/qqd_word_count
        qqd_uncertain_ratio = qqd_uncertain_count/qqd_word_count
        qqd_constrain_ratio = qqd_constrain_count/qqd_word_count
        mda_positive_ratio = mda_positive_count/mda_word_count
        if isinstance( mda_word_count, ( int) ) and isinstance( mda_negative_count, ( int) ):
            mda_negative_ratio = mda_negative_count/mda_word_count
        else:
            mda_negative_ratio = ""
        mda_uncertain_ratio = mda_uncertain_count/mda_word_count
        mda_constrain_ratio = mda_constrain_count/mda_word_count
        rf_positive_count = ""
        rf_positive_count = ""
        rf_negative_count = ""
        rf_constrain_count = ""
        rf_uncertain_count = ""
        rf_polarity = ""
        rf_word_count = ""
        rf_avg_sentence_lenght = ""
        rf_complex_words = ""
        rf_percentage_complex_words = ""
        rf_fog_idex = ""
        rf_fog_idex = ""
        rf_positive_ratio = ""
        rf_negative_ratio = ""
        rf_uncertain_ratio = ""
        rf_constrain_ratio = ""

    elif mda != 0 and qqd == 0 and rfs != 0:

        print(5)
        obj_clean_mda = cleanText(mda)
        obj_clean_rf = cleanText(rfs)
        obj_clean = obj_clean_mda.extend(obj_clean_rf)
        mda_positive_count = countPos(obj_clean_mda, positive)
        rf_positive_count = countPos(obj_clean_rf, positive)
        mda_negative_count = countNeg(obj_clean_mda, negative)
        rf_negative_count = countNeg(obj_clean_rf, negative)
        mda_constrain_count = countCons(obj_clean_mda, constrain)
        rf_constrain_count = countCons(obj_clean_rf, constrain)
        mda_uncertain_count = countUn(obj_clean_mda, uncertain)
        rf_uncertain_count = countUn(obj_clean_rf, uncertain)
        mda_polarity = getSentiment(obj_clean_mda, negative, positive)
        rf_polarity = getSentiment(obj_clean_rf, negative, positive)
        mda_word_count = len(obj_clean_mda)
        rf_word_count = len(obj_clean_rf)
        if rfs.count('.') == 0:
            rf_avg_sentence_lenght = 0
        else:
            rf_avg_sentence_lenght = len(rfs.split())/rfs.count('.')
        mda_avg_sentence_lenght = len(mda.split())/mda.count('.')
        mda_complex_words = complex_word(obj_clean_mda)
        rf_complex_words = complex_word(obj_clean_rf)
        mda_percentage_complex_words = (mda_complex_words/mda_word_count)*100
        
        mda_fog_idex = 0.4*(mda_avg_sentence_lenght +
                            mda_percentage_complex_words)
        
        mda_positive_ratio = mda_positive_count/mda_word_count
        mda_negative_ratio = mda_negative_count/mda_word_count
        mda_uncertain_ratio = mda_uncertain_count/mda_word_count
        mda_constrain_ratio = mda_constrain_count/mda_word_count
        
        if rf_word_count == 0:
            rf_percentage_complex_words = ""
            rf_positive_ratio = ""
            rf_negative_ratio = ""
            rf_uncertain_ratio = ""
            rf_constrain_ratio = ""
            rf_fog_idex = ""
        else:
            rf_percentage_complex_words = (rf_complex_words/rf_word_count)*100

            rf_positive_ratio = rf_positive_count/rf_word_count
            rf_negative_ratio = rf_negative_count/rf_word_count
            rf_uncertain_ratio = rf_uncertain_count/rf_word_count
            rf_constrain_ratio = rf_constrain_count/rf_word_count
            rf_fog_idex = ""

        qqd_positive_count = ""
        qqd_negative_count = ""
        qqd_constrain_count = ""
        qqd_uncertain_count = ""
        qqd_polarity = ""
        qqd_word_count = ""
        qqd_avg_sentence_lenght = ""
        qqd_complex_words = ""
        qqd_percentage_complex_words = ""
        qqd_fog_idex = ""
        qqd_positive_ratio = ""
        qqd_negative_ratio = ""
        qqd_uncertain_ratio = ""
        qqd_constrain_ratio = ""

    elif mda == 0 and qqd != 0 and rfs != 0:

        print(6)
        obj_clean_qqd = cleanText(qqd)
        obj_clean_rf = cleanText(rfs)
        obj_clean = obj_clean_qqd.extend(obj_clean_rf)
        qqd_positive_count = countPos(obj_clean_qqd, positive)
        rf_positive_count = countPos(obj_clean_rf, positive)
        qqd_negative_count = countNeg(obj_clean_qqd, negative)
        rf_negative_count = countNeg(obj_clean_rf, negative)
        qqd_constrain_count = countCons(obj_clean_qqd, constrain)
        rf_constrain_count = countCons(obj_clean_rf, constrain)
        qqd_uncertain_count = countUn(obj_clean_qqd, uncertain)
        rf_uncertain_count = countUn(obj_clean_rf, uncertain)
        qqd_polarity = getSentiment(obj_clean_qqd, negative, positive)
        rf_polarity = getSentiment(obj_clean_rf, negative, positive)
        qqd_word_count = len(obj_clean_qqd)
        rf_word_count = len(obj_clean_rf)
        if rfs.count('.') == 0:
            rf_avg_sentence_lenght = ""
        else:
            rf_avg_sentence_lenght = len(rfs.split())/rfs.count('.')
        qqd_avg_sentence_lenght = len(qqd.split())/qqd.count('.')
        qqd_complex_words = complex_word(obj_clean_qqd)
        rf_complex_words = complex_word(obj_clean_rf)
        qqd_percentage_complex_words = (qqd_complex_words/qqd_word_count)*100

        qqd_fog_idex = 0.4*(qqd_avg_sentence_lenght +
                            qqd_percentage_complex_words)
        if rf_word_count == 0:
            rf_fog_idex = ""
            rf_positive_ratio = ""
            rf_negative_ratio = ""
            rf_uncertain_ratio = ""
            rf_constrain_ratio = ""
            rf_percentage_complex_words = ""
        else:
            rf_fog_idex = ""
            rf_positive_ratio = rf_positive_count/rf_word_count
            rf_negative_ratio = rf_negative_count/rf_word_count
            rf_uncertain_ratio = rf_uncertain_count/rf_word_count
            rf_constrain_ratio = rf_constrain_count/rf_word_count
            rf_percentage_complex_words = (rf_complex_words/rf_word_count)*100
        qqd_positive_ratio = qqd_positive_count/qqd_word_count
        qqd_negative_ratio = qqd_negative_count/qqd_word_count
        qqd_uncertain_ratio = qqd_uncertain_count/qqd_word_count
        qqd_constrain_ratio = qqd_constrain_count/qqd_word_count
        mda_positive_count = ""
        mda_negative_count = ""
        mda_constrain_count = ""
        mda_uncertain_count = ""
        mda_polarity = ""
        mda_word_count = ""
        mda_avg_sentence_lenght = ""
        mda_complex_words = ""
        mda_percentage_complex_words = ""
        mda_fog_idex = ""
        mda_positive_ratio = ""
        mda_negative_ratio = ""
        mda_uncertain_ratio = ""
        mda_constrain_ratio = ""

    elif mda != 0 and qqd != 0 and rfs != 0:

        print(7)
        obj_clean_mda = cleanText(mda)
        obj_clean_qqd = cleanText(qqd)
        obj_clean_rf = cleanText(rfs)
        temp = obj_clean_qqd.extend(obj_clean_rf)
        
        mda_positive_count = countPos(obj_clean_mda, positive)
        qqd_positive_count = countPos(obj_clean_qqd, positive)
        rf_positive_count = countPos(obj_clean_rf, positive)
        mda_negative_count = countNeg(obj_clean_mda, negative)
        rf_negative_count = countNeg(obj_clean_rf, negative)
        qqd_negative_count = countNeg(obj_clean_qqd, negative)
        mda_constrain_count = countCons(obj_clean_mda, constrain)
        qqd_constrain_count = countCons(obj_clean_qqd, constrain)
        rf_constrain_count = countCons(obj_clean_rf, constrain)
        mda_uncertain_count = countUn(obj_clean_mda, uncertain)
        qqd_uncertain_count = countUn(obj_clean_qqd, uncertain)
        rf_uncertain_count = countUn(obj_clean_rf, uncertain)
        mda_polarity = getSentiment(obj_clean_mda, negative, positive)
        qqd_polarity = getSentiment(obj_clean_qqd, negative, positive)
        rf_polarity = getSentiment(obj_clean_rf, negative, positive)
        mda_word_count = len(obj_clean_mda)
        qqd_word_count = len(obj_clean_qqd)
        rf_word_count = len(obj_clean_rf)
        mda_avg_sentence_lenght = len(mda.split())/mda.count('.')
        qqd_avg_sentence_lenght = len(qqd.split())/qqd.count('.')
        if rfs.count('.') == 0:
            rf_avg_sentence_lenght = ""
        else:
            rf_avg_sentence_lenght = len(rfs.split())/rfs.count('.')
        if rf_word_count == 0:
            rf_fog_idex = ""
            rf_positive_ratio = ""
            rf_negative_ratio = ""
            rf_uncertain_ratio = ""
            rf_constrain_ratio = ""
            rf_percentage_complex_words = ""
        else:
            rf_fog_idex = ""
            rf_positive_ratio = rf_positive_count/rf_word_count
            rf_negative_ratio = rf_negative_count/rf_word_count
            rf_uncertain_ratio = rf_uncertain_count/rf_word_count
            rf_constrain_ratio = rf_constrain_count/rf_word_count
            rf_percentage_complex_words = ""
            
        mda_complex_words = complex_word(obj_clean_mda)
        qqd_complex_words = complex_word(obj_clean_qqd)
        rf_complex_words = complex_word(obj_clean_rf)
        mda_percentage_complex_words = (mda_complex_words/mda_word_count)*100
        qqd_percentage_complex_words = (qqd_complex_words/qqd_word_count)*100
        
        mda_fog_idex = 0.4*(mda_avg_sentence_lenght +
                            mda_percentage_complex_words)
        qqd_fog_idex = 0.4*(qqd_avg_sentence_lenght +
                            qqd_percentage_complex_words)
        
        qqd_positive_ratio = qqd_positive_count/qqd_word_count
        qqd_negative_ratio = qqd_negative_count/qqd_word_count
        qqd_uncertain_ratio = qqd_uncertain_count/qqd_word_count
        qqd_constrain_ratio = qqd_constrain_count/qqd_word_count
        mda_positive_ratio = mda_positive_count/mda_word_count
        mda_negative_ratio = mda_negative_count/mda_word_count
        mda_uncertain_ratio = mda_uncertain_count/mda_word_count
        mda_constrain_ratio = mda_constrain_count/mda_word_count
    else:
        mda_positive_count = ""
        mda_negative_count = ""
        mda_constrain_count = ""
        mda_uncertain_count = ""
        mda_polarity = ""
        mda_word_count = ""
        mda_avg_sentence_lenght = ""
        mda_complex_words = ""
        mda_percentage_complex_words = ""
        mda_fog_idex = ""
        mda_positive_ratio = ""
        mda_negative_ratio = ""
        mda_uncertain_ratio = ""
        mda_constrain_ratio = ""
        qqd_positive_count = ""
        qqd_negative_count = ""
        qqd_constrain_count = ""
        qqd_uncertain_count = ""
        qqd_polarity = ""
        qqd_word_count = ""
        qqd_avg_sentence_lenght = ""
        qqd_complex_words = ""
        qqd_percentage_complex_words = ""
        qqd_fog_idex = ""
        qqd_positive_ratio = ""
        qqd_negative_ratio = ""
        qqd_uncertain_ratio = ""
        qqd_constrain_ratio = ""
        rf_positive_count = ""
        rf_negative_count = ""
        rf_constrain_count = ""
        rf_uncertain_count = ""
        rf_polarity = ""
        rf_word_count = ""
        rf_avg_sentence_lenght = ""
        rf_complex_words = ""
        rf_percentage_complex_words = ""
        rf_fog_idex = ""
        rf_positive_ratio = ""
        rf_negative_ratio = ""
        rf_uncertain_ratio = ""
        rf_constrain_ratio = ""
    """ writing on to excel """
    sheet1.write(i, 6, mda_positive_count)
    sheet1.write(i, 7, mda_negative_count)
    sheet1.write(i, 8, mda_avg_sentence_lenght)
    sheet1.write(i, 9, mda_percentage_complex_words)
    sheet1.write(i, 10, mda_fog_idex)
    sheet1.write(i, 11, mda_complex_words)
    sheet1.write(i, 12, mda_word_count)
    sheet1.write(i, 13, mda_uncertain_count)
    sheet1.write(i, 14, mda_constrain_count)
    sheet1.write(i, 15, mda_positive_ratio)
    sheet1.write(i, 16, mda_negative_ratio)
    sheet1.write(i, 17, mda_constrain_ratio)
    sheet1.write(i, 18, qqd_positive_count)
    sheet1.write(i, 19, qqd_negative_count)
    sheet1.write(i, 20, qqd_avg_sentence_lenght)
    sheet1.write(i, 21, qqd_percentage_complex_words)
    sheet1.write(i, 22, qqd_fog_idex)
    sheet1.write(i, 23, qqd_complex_words)
    sheet1.write(i, 24, qqd_word_count)
    sheet1.write(i, 25, qqd_uncertain_count)
    sheet1.write(i, 26, qqd_constrain_count)
    sheet1.write(i, 27, qqd_positive_ratio)
    sheet1.write(i, 28, qqd_negative_ratio)
    sheet1.write(i, 29, qqd_constrain_ratio)
    sheet1.write(i, 30, rf_positive_count)
    sheet1.write(i, 31, rf_negative_count)
    sheet1.write(i, 32, rf_avg_sentence_lenght)
    sheet1.write(i, 33, rf_percentage_complex_words)
    sheet1.write(i, 34, rf_fog_idex)
    sheet1.write(i, 35, rf_complex_words)
    sheet1.write(i, 36, rf_word_count)
    sheet1.write(i, 37, rf_uncertain_count)
    sheet1.write(i, 38, rf_constrain_count)
    sheet1.write(i, 39, rf_positive_ratio)
    sheet1.write(i, 41, rf_negative_ratio)
    sheet1.write(i, 42, rf_constrain_ratio)
    sheet1.write(i, 43, total_cons)
workbook.close()
