import xlrd
from xlrd import xlsx
import xlwt
import os
import xlwt
from xlwt import *
import nltk
from nltk import tokenize
import re
from nltk.tokenize import WordPunctTokenizer
import numpy as np

###List of Convergent
Convergent_words_search = ["Which", "which", "What", "what", "Do", "do", "Does", "does", "Is", "is", "Are", "are",
                           "Has", "has", "have", "Have", "Who", "who"
                                                                "whom", "Whom"
                                                                        "How", "how", "When", "when", "Why", "why"]
ID_Convergent_questions = []
essay_ID_Convergent_questions = []
comment_Convergent_questions = []
gold_dir_Convergent_questions = []
gold_rev_succ10_Convergent_questions = []
original_Convergent_questions = []
revised_Convergent_questions = []
original_plus_Convergent_questions = []
revised_plus_Convergent_questions = []
###List of Divergent
Divergent_words_search = ["Could", "such", "Such", "Can", "could", "can", "If", "if"]
ID_Divergent_questions = []
essay_ID_Divergent_questions = []
comment_Divergent_questions = []
gold_dir_Divergent_questions = []
gold_rev_succ10_Divergent_questions = []
original_Divergent_questions = []
revised_Divergent_questions = []
original_plus_Divergent_questions = []
revised_plus_Divergent_questions = []
###List of None questions
ID_None_questions = []
essay_ID_None_questions = []
comment_None_questions = []
gold_dir_None_questions = []
gold_rev_succ10_None_questions = []
original_None_questions = []
revised_None_questions = []
original_plus_None_questions = []
revised_plus_None_questions = []


def data_preprocess():
    Convergent_words_search = ["Which",  "What", "Do",  "Does",  "Is", "Are",
                               "Has", "Have", "Who",  "Whom", "How", "When",
                               "Why"]
    Divergent_words_search = ["Could", "Can", "If"]

    Question_sign_search = ["?"]

    data = xlrd.open_workbook("Corpus_data2.xlsx")
    sheet1 = data.sheet_by_name('Corpus')
    ID = sheet1.col_values(0)
    ID.pop(0)
    essay_ID = sheet1.col_values(1)
    essay_ID.pop(0)
    comment = sheet1.col_values(3)  # classify
    comment.pop(0)
    gold_dir = sheet1.col_values(6)
    gold_dir.pop(0)
    gold_rev_succ10 = sheet1.col_values(9)
    gold_rev_succ10.pop(0)
    original = sheet1.col_values(10)
    original.pop(0)
    revised = sheet1.col_values(11)
    revised.pop(0)
    original_plus = sheet1.col_values(12)
    original_plus.pop(0)
    revised_plus = sheet1.col_values(13)
    revised_plus.pop(0)
    i = -1
    for single_sentence in comment:
        # sen_tokenizer = nltk.data.load('tokenizers/punkt/english.pickle')
        sentence_seperate = nltk.word_tokenize(single_sentence)
        i += 1
        for words in sentence_seperate:
            if words in Convergent_words_search:
                #print(words)
                comment_Convergent_questions.append(single_sentence)
                #b = comment.index(single_sentence)
                ID_Convergent_questions.append(ID[i])
                essay_ID_Convergent_questions.append(essay_ID[i])
                gold_dir_Convergent_questions.append(gold_dir[i])
                gold_rev_succ10_Convergent_questions.append(gold_rev_succ10[i])
                original_Convergent_questions.append(original[i])
                revised_Convergent_questions.append(revised[i])
                original_plus_Convergent_questions.append(original_plus[i])
                revised_plus_Convergent_questions.append(revised_plus[i])
                break
            elif words in Divergent_words_search:
                comment_Divergent_questions.append(single_sentence)
                #c = comment.index(single_sentence)
                ID_Divergent_questions.append(ID[i])
                essay_ID_Divergent_questions.append(essay_ID[i])
                gold_dir_Divergent_questions.append(gold_dir[i])
                gold_rev_succ10_Divergent_questions.append(gold_rev_succ10[i])
                original_Divergent_questions.append(original[i])
                revised_Divergent_questions.append(revised[i])
                original_plus_Divergent_questions.append(original_plus[i])
                revised_plus_Divergent_questions.append(revised_plus[i])
                break
            elif
        else:
            #     # elif words not in Convergent_words_search and words not in Divergent_words_search:
            comment_None_questions.append(single_sentence)
            d = comment.index(single_sentence)
            ID_None_questions.append(ID[i])
            essay_ID_None_questions.append(essay_ID[i])
            gold_dir_None_questions.append(gold_dir[i])
            gold_rev_succ10_None_questions.append(gold_rev_succ10[i])
            original_None_questions.append(original[i])
            revised_None_questions.append(revised[i])
            original_plus_None_questions.append(original_plus[i])
            revised_plus_None_questions.append(revised_plus[i])


# print(comment_None_questions)

# print("This is convergent questions:", comment_Convergent_questions)
# print("This is divergent questions:", comment_Divergent_questions)
# print("This is none question:", comment_None_questions)


def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


# ### writing into the excel file
def write_excel():
    f = xlwt.Workbook(encoding='utf-8')
    sheet1 = f.add_sheet('Convergent questions', cell_overwrite_ok=True)
    sheet2 = f.add_sheet('Divergent questions', cell_overwrite_ok=True)
    sheet3 = f.add_sheet('None questions', cell_overwrite_ok=True)

    row0 = ["ID", "essay ID", "comment", "gold_dir", "gold_rev_succ10", "original", "revised", "original_plus",
            "revised_plus"]
    columA0 = ID_Convergent_questions
    columA1 = essay_ID_Convergent_questions
    columA2 = comment_Convergent_questions
    columA3 = gold_dir_Convergent_questions
    columA4 = gold_rev_succ10_Convergent_questions
    columA5 = original_Convergent_questions
    columA6 = revised_Convergent_questions
    columA7 = original_plus_Convergent_questions
    columA8 = revised_plus_Convergent_questions

    columB0 = ID_Divergent_questions
    columB1 = essay_ID_Divergent_questions
    columB2 = comment_Divergent_questions
    columB3 = gold_dir_Divergent_questions
    columB4 = gold_rev_succ10_Divergent_questions
    columB5 = original_Divergent_questions
    columB6 = revised_Divergent_questions
    columB7 = original_plus_Divergent_questions
    columB8 = revised_plus_Divergent_questions

    columC0 = ID_None_questions
    columC1 = essay_ID_None_questions
    columC2 = comment_None_questions
    columC3 = gold_dir_None_questions
    columC4 = gold_rev_succ10_None_questions
    columC5 = original_None_questions
    columC6 = revised_None_questions
    columC7 = original_plus_None_questions
    columC8 = revised_plus_None_questions

    # style = xlwt.XFStyle()
    default = set_style('Times New Roman', 220, True)

    for c in range(0, len(row0)):
        # sheet1.col(c).width = 8888
        sheet1.write(0, c, row0[c], set_style('Times New Roman', 220, True))
        sheet2.write(0, c, row0[c], set_style('Times New Roman', 220, True))
        sheet3.write(0, c, row0[c], set_style('Times New Roman', 220, True))

    for j in range(0, len(columA0)):
        # sheet1.col(c).width = 256 * 20
        sheet1.write(j + 1, 0, columA0[j], default)
        sheet1.write(j + 1, 1, columA1[j], default)
        sheet1.write(j + 1, 2, columA2[j], default)
        sheet1.write(j + 1, 3, columA3[j], default)
        sheet1.write(j + 1, 4, columA4[j], default)
        sheet1.write(j + 1, 5, columA5[j], default)
        sheet1.write(j + 1, 6, columA6[j], default)
        sheet1.write(j + 1, 7, columA7[j], default)
        sheet1.write(j + 1, 8, columA8[j], default)
    for k in range(0, len(columB0)):
        # sheet2.col(c).width = 256 * 20
        sheet2.write(k + 1, 0, columB0[k], default)
        sheet2.write(k + 1, 1, columB1[k], default)
        sheet2.write(k + 1, 2, columB2[k], default)
        sheet2.write(k + 1, 3, columB3[k], default)
        sheet2.write(k + 1, 4, columB4[k], default)
        sheet2.write(k + 1, 5, columB5[k], default)
        sheet2.write(k + 1, 6, columB6[k], default)
        sheet2.write(k + 1, 7, columB7[k], default)
        sheet2.write(k + 1, 8, columB8[k], default)
    for m in range(0, len(columC0)):
        # sheet3.col(c).width = 256 * 20
        sheet3.write(m + 1, 0, columC0[m], default)
        sheet3.write(m + 1, 1, columC1[m], default)
        sheet3.write(m + 1, 2, columC2[m], default)
        sheet3.write(m + 1, 3, columC3[m], default)
        sheet3.write(m + 1, 4, columC4[m], default)
        sheet3.write(m + 1, 5, columC5[m], default)
        sheet3.write(m + 1, 6, columC6[m], default)
        sheet3.write(m + 1, 7, columC7[m], default)
        sheet3.write(m + 1, 8, columC8[m], default)
    f.save('Cityu.xls')


if __name__ == "__main__":
    data_preprocess()
    write_excel()
