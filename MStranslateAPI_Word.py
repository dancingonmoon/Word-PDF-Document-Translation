import json
import uuid

import docx
import requests
import math
import re
# import copy

# define the function, input the text and ouput the translation:
def MStranslation_API(
    text,
    lang_in="en",
    lang_out="zh-Hans",
    subscription_key="pls input your MStranslation API key",
):
    """
    To translate the pre-defined text:
    args:
        text: text or text list, to be translated;
        lang_in : language input;
        lang_out: language output, language list is accepted, for example:['ru','zh-Hans']
        subscription_key: Microsoft Translation API Key;
    out:
        trans_text: the tranlated text list (when single language output), or translated text dictionary (when multiple language output); when in dictionary, the key is the language names, the value is the list of translated text
        response: API output, as backup, in case of overseeing the errors;
    """

    # Add your subscription key and endpoint
    subscription_key = subscription_key
    endpoint = "https://api.cognitive.microsofttranslator.com"  
    # Add your location, also known as region. The default is global.
    # This is required if using a Cognitive Services resource.
    location = "eastasia"
    path = "/translate"
    constructed_url = endpoint + path

    params = {"api-version": "3.0", "from": lang_in, "to": lang_out}

    headers = {
        "Ocp-Apim-Subscription-Key": subscription_key,
        "Ocp-Apim-Subscription-Region": location,
        "Content-type": "application/json",
        "X-ClientTraceId": str(uuid.uuid4()),
    }

    # You can pass more than one object in body.
    body = []

    if isinstance(text, list):
        for txt in text:
            body.append({"text": txt})

    if isinstance(text, str):
        body.append({"text": text})

    # print(body)
    # to turn lang_out into list, when lang_out is with one language:
    if isinstance(lang_out, str):
        lang_out = [lang_out]

    request = requests.post(constructed_url, params=params, headers=headers, json=body)
    response = request.json()
    # response_json = json.dumps(response, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ': '))
    # print(response_json)

    trans_text = []

    try:
        if isinstance(lang_out, list) and len(lang_out) > 1:
            trans_text = {}
            for i, lang in enumerate(lang_out):
                tmp = []
                for r in response:
                    tmp.append(r["translations"][i]["text"])
                trans_text[lang] = tmp
        else:
            for r in response:
                trans_text.append(r["translations"][0]["text"])
    except:
        print(response)

    return trans_text, response

def MStranslation_dynamicDictionary_API(
    text,
    dynamic_dict=False,
    lang_in="en",
    lang_out="zh-Hans",
    subscription_key="pls input your MStranslation API key",
):
    """
    Microsoft Translation API with Dynamic dictionary:
    args:
        text: text or text list to be translated;
        dynamic_dict: dynamic dictionary, comprises of specialized vocabulary, production name, personal names etc., that has conventional translated phrases for example: {'莫言':'Mr.Moyan'};
        lang_in : language input;
        lang_out: language output, list is accepted to add multiple language, for example:['ru','zh-Hans']
        subscription_key: Microsoft Translation API key;
    out:
        trans_text: the text list (when single language output), or text dictionary (when multiple language), dictionary shall have language name as the key, and translated text list as value
        response: API output, as backup, in case of overseeing the errors;
    """

    # Add your subscription key and endpoint
    subscription_key = subscription_key
    endpoint = "https://api.cognitive.microsofttranslator.com"  
    # Add your location, also known as region. The default is global.
    # This is required if using a Cognitive Services resource.
    location = "eastasia"
    path = "/translate"
    constructed_url = endpoint + path

    params = {"api-version": "3.0", "from": lang_in, "to": lang_out}

    headers = {
        "Ocp-Apim-Subscription-Key": subscription_key,
        "Ocp-Apim-Subscription-Region": location,
        "Content-type": "application/json",
        "X-ClientTraceId": str(uuid.uuid4()),
    }

    # You can pass more than one object in body.
    body = []

    if isinstance(text, list):
        for txt in text:
            if isinstance(dynamic_dict, dict):
                for key in dynamic_dict.keys():
                    # if txt comprises of dyanamic_dict key, then key shall be replaced by value.
                    sub_txt = (
                        '<mstrans:dictionary translation="'
                        + dynamic_dict[key]
                        + '">'
                        + key
                        + "</mstrans:dictionary>"
                    )
                    txt = re.sub(key, sub_txt, txt)
                body.append({"text": txt})
            elif dynamic_dict == False or dynamic_dict == "":
                body.append({"text": txt})
            else:
                print("Neither False nor Dictionary dynamic_dict is !")

    if isinstance(text, str):  # when text is one single text string
        if isinstance(dynamic_dict, dict):
            for key in dynamic_dict.keys():
                # if txt comprises of dyanamic_dict key, then key shall be replaced by value.
                sub_txt = (
                    '<mstrans:dictionary translation="'
                    + dynamic_dict[key]
                    + '">'
                    + key
                    + "</mstrans:dictionary>"
                )
                text = re.sub(key, sub_txt, text)
            body.append({"text": text})
        elif dynamic_dict == False or dynamic_dict == "":
            body.append({"text": text})
        else:
            print("Neither False nor Dictionary dynamic_dict is !")

    # when lang_out is with single language, to turn lang_out into list:
    if isinstance(lang_out, str):
        lang_out = [lang_out]

    request = requests.post(constructed_url, params=params, headers=headers, json=body)
    response = request.json()
    # response_json = json.dumps(response, sort_keys=True, ensure_ascii=False, indent=4, separators=(',', ': '))
    # print(response_json)

    trans_text = []

    try:
        if isinstance(lang_out, list) and len(lang_out) > 1:
            trans_text = {}
            for i, lang in enumerate(lang_out):
                tmp = []
                for r in response:
                    tmp.append(r["translations"][i]["text"])
                trans_text[lang] = tmp
        else:
            for r in response:
                trans_text.append(r["translations"][0]["text"])
    except:
        print(response)

    return trans_text, response



# to define fucntion, for each 'Run' in each paragraph object (stalled in memory), perceive whether the "Run" contains inlined picture/drawing, only non-null text shall be applied and subsititued:
def paragraph_runs_replace(
    paragraph,
    text,
):
    """
    To replace the text of (inside) each run of each paragraph, with inlined shape/picture/drawings untouched;
    1. for the paragraph object, go through each run, get the ratio between run.text and paragraph.text, then apply such ratio upon translated text, so as to distribute the translated text upon each of "Run" and get the distriubted run.text start and stop pointer in the form of tuple;
    2. for each run, to replace each of run.text , when run.text is non-null.
    Note: the fuction revoke the library of python-docx, which operate on document object. Document object is stalled in memory; simple copy or deep copy can't work on object.
    args:
        paragraph: paragraph object, could be from document, or table, or header,footer etc.;
        text: replaced text, single string, not list;
    out:
        paragraph: paragraph object that complete the operation;
    """
    text_len = len(text)  # to get the length of text string;
    source_len = len(paragraph.text)

    run_attr = {}
    pointer = 0
    for i, r in enumerate(paragraph.runs):
        if source_len == 0:
            run_attr[i] = 0
        else:
            len_distrib = math.ceil(len(r.text) / source_len * text_len)  # round up to integer
            run_attr[i] = (pointer, pointer + len_distrib)
            pointer = pointer + len_distrib
    # print(run_attr)
    
    # go though each run, replace run.text if non-null:
    if len(paragraph.runs) != 0:
        for i, r in enumerate(paragraph.runs):
            if r.text != "":
                # replace run.text, with style and font unchanged:
                r.text = text[run_attr[i][0] : run_attr[i][1]]

    return paragraph

def Word_MStranslation(
    doc,
    dynamic_dict=False,
    lang_in='zh-Hans',
    lang_out='en',    
    filename=False,
    subscription_key="pls input your MStranslation API key",
):
    """
    To work on paragraph, table, header and footer, and replace those text with translated text, and style and font remain; if dynamic dictionary is in need, text with dictionary keys shall be replaced with dictionary value;
    A: paragraph:
    1. generate the dictionary for the paragraphs structure, and the text;
    2. Set up dictionary, with the key has the value that is the index No. inside the test list;
    2. invoke the translation API, to generate translted_text list;
    3. for each paragraph, to replace the text with translated text, with style and font unchanged;
    B: Table :
    1. generate the dictionary for the table structure, and the text
    2. Set up dictionary, with the key has the value that is the index No. inside the test list;
    2. invoke the translation API, to generate translted_text list;
    3. for each table, to replace the text with translated text, with style and font unchanged;
    C: header and footer translation;
    args:
        doc: document object generated from python-docx;
        lang_in : language input;
        lang_out: language output, in single language;
        dynamic_dict: dynamic dictionary, comprises of specialized vocabulary, production name, personal names etc., that has conventional translated phrases for example: {'莫言':'Mr.Moyan'};
        filename: whether to save into specified file path; default value:  False, otherwise, the file path
        subscription_key: Microsoft API key;
    out:
        doc: the document object after the operation;
        when filename!=False, to save document into specified path;
        if error comes out, API outputs the response with error code;
    """
    # Paragraphs Translation:
    # -----------------------------------
    # generate text list to be translated
    text_dict = {}
    for i, para in enumerate(doc.paragraphs):
        text_dict[(i)] = para.text
    # print(text_dict)
    text = list(text_dict.values())
    # Set up dictionary, with the key has the value that is the index No. inside the test list;
    textindex_dict = {}
    for key in text_dict.keys():
        textindex_dict[key] = text.index(text_dict[key])

    # generate the translated text list
    # invoke API, to deicde which API is gona to be used ,subject to the dynamic dictionary;
    if dynamic_dict == False:
        trans_text, _ = MStranslation_API(
            text, lang_in=lang_in, lang_out=lang_out, subscription_key=subscription_key
        )
    elif isinstance(dynamic_dict, dict):
        trans_text, _ = MStranslation_dynamicDictionary_API(
            text,
            dynamic_dict=dynamic_dict,
            lang_in=lang_in,
            lang_out=lang_out,
            subscription_key=subscription_key,
        )
    else:
        print("Neither False nor Dictionary dynamic_dict is !")
        
    # for each paragraph:
    for i, para in enumerate(doc.paragraphs):
        para_trans_text = trans_text[textindex_dict[(i)]]
        paragraph_runs_replace(para, para_trans_text)

    # -------------------------------------------

    # Table translation:
    # -------------------------------------
    # get the dictionary for the table structure, and the text
    text_dict = {}
    for t, table in enumerate(doc.tables):
        for r, row in enumerate(table.rows):
            for c, cell in enumerate(row.cells):
                for p, para in enumerate(cell.paragraphs):
                    text_dict[(t, r, c, p)] = para.text

    # print('text字典len:{}'.format(len(text)))
    text = list(text_dict.values())
    # Set up dictionary, with the key has the value that is the index No. inside the test list;
    textindex_dict = {}
    for key in text_dict.keys():
        textindex_dict[key] = text.index(text_dict[key])
        
    # invoke API, to deicde which API is gona to be used ,subject to the dynamic dictionary;
    if dynamic_dict == False:
        trans_text, _ = MStranslation_API(
            text, lang_in=lang_in, lang_out=lang_out, subscription_key=subscription_key
        )
    elif isinstance(dynamic_dict, dict):
        trans_text, _ = MStranslation_dynamicDictionary_API(
            text,
            dynamic_dict=dynamic_dict,
            lang_in=lang_in,
            lang_out=lang_out,
            subscription_key=subscription_key,
        )
    else:
        print("Neither False nor Dictionary dynamic_dict is !")

    # for each paragraph of table structure:
    for t, table in enumerate(doc.tables):
        for r, row in enumerate(table.rows):
            for c, cell in enumerate(row.cells):
                for p, para in enumerate(cell.paragraphs):

                    para_trans_text = trans_text[textindex_dict[(t, r, c, p)]]
                    paragraph_runs_replace(para, para_trans_text)

    # ----------------------------------
    # header, footer , translation:
    # ----------------------------------
    # to get all section.header.paragraphs和secction.footer.paragraphs structure in dictionary and the text
    text_dict = {}
    for s, section in enumerate(doc.sections):
        for p, para in enumerate(section.header.paragraphs):  # each section has only one header;
            text_dict[(s, "header", p)] = para.text
        for p, para in enumerate(section.footer.paragraphs):  # each section only has one footer;
            if para.text.isdigit() == False and para.text != "":  # footer that has dynamic page No, let it be
                text_dict[(s, "footer", p)] = para.text
    # print(text_dict)
    text = list(text_dict.values())
    # Set up dictionary, with the key has the value that is the index No. inside the test list;   
    textindex_dict = {}
    for key in text_dict.keys():
        textindex_dict[key] = text.index(text_dict[key])

    # invoke API, to deicde which API is gona to be used ,subject to the dynamic dictionary;
    if dynamic_dict == False:
        trans_text, _ = MStranslation_API(
            text, lang_in=lang_in, lang_out=lang_out, subscription_key=subscription_key
        )
    elif isinstance(dynamic_dict, dict):
        trans_text, _ = MStranslation_dynamicDictionary_API(
            text,
            dynamic_dict=dynamic_dict,
            lang_in=lang_in,
            lang_out=lang_out,
            subscription_key=subscription_key,
        )
    else:
        print("Neither False nor Dictionary dynamic_dict is !")
        
    # for each paragraph of all section.header.paragraphs:
    for s, section in enumerate(doc.sections):
        for p, para in enumerate(section.header.paragraphs):  # each section only has one header;
            para_trans_text = trans_text[textindex_dict[(s, "header", p)]]
            paragraph_runs_replace(para, para_trans_text)
        for p, para in enumerate(section.footer.paragraphs):  # each section only has one header;
            if para.text.isdigit() == False and para.text != "":  # footer with dynamic pages No, leave it unchanged;
                para_trans_text = trans_text[textindex_dict[(s, "footer", p)]]
                paragraph_runs_replace(para, para_trans_text)
    # ---------------------------------
    # save,output:
    if filename != False:
        doc.save(filename)
    return doc



