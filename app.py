import streamlit as st
import string
import pandas as pd
import urllib.parse, urllib.request
from urllib import parse
from youtube_transcript_api import YouTubeTranscriptApi
import json
import spacy
from spacytextblob.spacytextblob import SpacyTextBlob
import re
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import queue
import os
# 영어 영상인데도 안되는 경우: en_core_web_sm 오류
# python -m spacy download en

#key
userKey = st.secrets.userKey
youtube_key = st.secrets.youtube_key

class Script_Exctractor:
    def __init__(self,vid,setTime = 300.0):
        ###
        #   vid: youtube url
        #   fileName: 저장 할 파일이름
        #   setTime: segmentation할 시간 단위
        #   기본값 300sec(5분)
        ###
        self.vid = vid
        self.scriptData:list = []
        self.setTime = setTime

    ### youtube script 추출
    def Extract(self):
        ###
        #   youtube script 받아오기 위한 url 전처리부분
        ###
        parsedUrl = parse.urlparse(self.vid)
        vid = parse.parse_qs(parsedUrl.query)['v'][0]
        languages = ['en','en-US']
        str:list = YouTubeTranscriptApi.get_transcript(vid,languages)

        ###
        #   저장된 json정보를 timeSet단위에 맞게 분리
        ###
        ret = queue.Queue()
        nowSec = self.setTime
        sentence = ''
        for st in str:
            if st['start'] >= nowSec:
                ret.put(sentence)
                sentence = ''
                nowSec += self.setTime
                if frontSt['start'] + frontSt['duration'] >= nowSec:
                    sentence += frontSt['text'] + '\n'
            sentence += st['text'] + '\n'
            frontSt = st
        ret.put(sentence)
        while not ret.empty():
            self.scriptData.append(ret.get())

        for i in range(len(self.scriptData)):
            text = self.scriptData[i].replace(u'\xa0', u' ').replace(u'\n',u' ').replace(u'  ',u' ')
            self.scriptData[i] = text


def extract_video_id(youtube_link):
    # 정규표현식으로 영상 ID 추출
    video_id = re.findall(r'v=([^\&]+)', youtube_link)
    return video_id[0] if video_id else None

def remove_special_characters(input_string):
    # string 모듈의 punctuation 상수에 있는 모든 특수 기호들을 빈 문자열로 치환하여 제거합니다.
    # 만약, 영문자와 숫자만 남기고 나머지를 제거하고 싶다면, string.ascii_letters + string.digits 를 사용하면 됩니다.
    return re.sub('[^a-zA-Z0-9]',' ',input_string).strip()

def get_youtube_video_title(youtube_link, api_key):
    # API 초기화
    youtube = build('youtube', 'v3', developerKey=api_key)

    # 유튜브 영상 ID 파싱
    video_id = extract_video_id(youtube_link)

    if youtube_link == "":
        return None
    elif not video_id:
        print("유효한 유튜브 링크 형식이 아닙니다.")
        return None

    try:
        # 영상 정보 요청
        video_response = youtube.videos().list(
            part='snippet',
            id=video_id
        ).execute()

        # 영상 제목 가져오기
        video_title = video_response['items'][0]['snippet']['title']
        print("get_youtube_video: "+video_title)
        video_title = remove_special_characters(video_title)
        #video_title = video_title.replace('?', u'').replace('!',u'').replace('\\',u'').replace('\\\\',)
        return video_title
    except HttpError as e:
        print(f"An HTTP error occurred: {e}")
        return None

def CallWikifier(text, lang, threshold, numberOfKCs):
    # Prepare the URL.
    data = urllib.parse.urlencode([
        ("text", text), ("lang", lang),
        ("userKey", userKey),
        ("pageRankSqThreshold", "%g" % threshold),
        ("applyPageRankSqThreshold", "true"),
        ("nTopDfValuesToIgnore", "200"),
        ("nWordsToIgnoreFromList", "200"),
        ("wikiDataClasses", "false"),
        ("wikiDataClassIds", "false"),
        ("support", "false"),
        ("ranges", "false"),
        ("minLinkFrequency", "3"),
        ("includeCosines", "false"),
        ("maxMentionEntropy", "2")
        ])
    url = "http://www.wikifier.org/annotate-article"
    # Call the Wikifier and read the response.
    req = urllib.request.Request(url, data=data.encode("utf8"), method="POST")
    with urllib.request.urlopen(req, timeout = 60) as f:
        response = f.read()
        response = json.loads(response.decode("utf8"))

    sorted_data = sorted(response['annotations'], key=lambda x: x['pageRank'], reverse=True)
    # Output the annotations.
    num = 0
    result = []
    for annotation in sorted_data:
        if num < numberOfKCs:
            print("%s (%s) %s" % (annotation["title"], annotation["url"], annotation['pageRank']))
            result.append({"title":annotation["title"],"url":annotation["url"],"pageRank":annotation["pageRank"]})

        num += 1

    res = result
    result = []
    return res

def combine_csv_to_excel(pr_csv_files, sa_csv_files, output_excel_file, sheet_names):
    writer = pd.ExcelWriter(output_excel_file, engine='xlsxwriter')

    for i in range(len(sheet_names)):
        pr_df = pd.read_csv(pr_csv_files[i])
        sa_df = pd.read_csv(sa_csv_files[i])
        sheet_name = sheet_names[i] if i < len(sheet_names) else f'Sheet{i+1}'
        if len(sheet_name) > 20:
            sheet_name = sheet_name[:20]  # 시트 이름을 20자 이후는 잘라냅니다.
        pr_df.to_excel(writer, sheet_name=sheet_name, index=False)
        sa_df.to_excel(writer, sheet_name=sheet_name+"_sentiment", index=False)

    writer._save()
    print("CSV 파일이 성공적으로 엑셀 파일로 합쳐졌습니다.")

def one_url_to_csv(want_url, want_time, want_num):
    video_title = get_youtube_video_title(want_url, youtube_key)
    print(video_title)
    Scripts = Script_Exctractor(want_url, want_time)
    Scripts.Extract()

    sentiments = Spacytextblob([Scripts.scriptData])
    sentiment_df = pd.DataFrame(sentiments.data)

    # segment counts
    number = 1
    results = []
    for text in Scripts.scriptData:
        print(f"{number}st segemnt")
        results.append(CallWikifier(text,"en",0.8,want_num))
        number += 1

    v_data = pd.DataFrame()
    seg_no = 1

    for seg_item in results:
        seg_index = range(0,len(seg_item))
        seg_df = pd.DataFrame(seg_item,index = seg_index)
        seg_df['seg_no'] = seg_no
        v_data = pd.concat([v_data,seg_df])
        seg_no = seg_no + 1

    pr_csv_filename = video_title+'.csv' #페이지랭크 파일이름
    sa_csv_filename = video_title+'_sentiment_analisys.csv' # 감성분석 파일 이름
    v_data.to_csv(pr_csv_filename)
    sentiment_df.to_csv(sa_csv_filename)
    
    return pr_csv_filename,sa_csv_filename, video_title, Scripts.scriptData

def delete_file(filename):
    try:
        os.remove(filename)
        print(f"{filename} 파일이 삭제되었습니다.")
    except FileNotFoundError:
        print(f"{filename} 파일을 찾을 수 없습니다.")
    except PermissionError:
        print(f"{filename} 파일에 대한 삭제 권한이 없습니다.")
    except Exception as e:
        print(f"파일 삭제 중 에러가 발생했습니다: {e}")

# 감성 분석
class Spacytextblob:
    def __init__(self, script_list):
        self.script_list = script_list
        self.data = []
        self.nlp = spacy.load('en_core_web_sm')
        self.nlp.add_pipe('spacytextblob')
        for segment in self.script_list:
            self.spacytextblob_print(segment)

    def spacytextblob_print(self, segment):
        doc_list=[]
        for text in segment:
            doc_list.append(self.nlp(text))

        idx=0
        sentiments_list = []
        for doc in doc_list:
            data = {"polarity":doc._.blob.polarity,"subjectivity":doc._.blob.subjectivity,"sentiment_assesments":doc._.blob.sentiment_assessments.assessments}
            sentiments_list.append(data)
            print(f"{idx+1}st")
            print(f"polarity:{doc._.blob.polarity}")
            print(f"subjectivity: {doc._.blob.subjectivity}")
            print(f"sentiment_assesments: {str(doc._.blob.sentiment_assessments.assessments)[:100]}...") #길이 100만큼만 출력

            idx += 1
        self.data = sentiments_list

def Analysis(split_seconds, num_concepts, url_list):
    pr_csv_list=[] #페이지랭크 파일이름 리스트
    sa_csv_list=[] #감성분석 파일이름 리스트
    title_list=[]
    want_url_list = url_list
    script_list=[]

    want_time = split_seconds
    want_num = num_concepts

    for want_url in want_url_list:
        try:
            pr_csv_filename,sa_csv_filename, title, scriptData = one_url_to_csv(want_url, want_time, want_num)
            pr_csv_list.append(pr_csv_filename)
            sa_csv_list.append(sa_csv_filename)
            title_list.append(title)
            script_list.append(scriptData)
        except:
            if want_url == "":
                pass
            else:
                print(want_url+" 이 영상은 변환에 실패했습니다.\n1. 영어 영상이 아닐 수 있습니다.\n2. API 서버와 통신이 안될 수 있습니다.(특히 wikifier)")


    output_excel_file = "combined_excel.xlsx"  # 생성할 엑셀 파일명을 지정해주세요.

    combine_csv_to_excel(pr_csv_list, sa_csv_list, output_excel_file, title_list)

    # 파일 삭제 함수 호출
    for title in title_list:
        delete_file(title+".csv")
        delete_file(title+"_sentiment_analisys.csv")

if __name__ == "__main__":
    # 제목
    st.title("동영상 감성 분석 웹사이트")

    # 내용
    st.write("이 웹사이트는 동영상 감성 분석을 위한 도구입니다. 아래에서 URL을 입력하고 제출하세요.")

    # 사이드 바에 설정 옵션 추가
    with st.sidebar:
        st.header("설정 옵션")
        split_seconds = st.number_input("영상을 몇 초로 나눌지 입력", value=300, key="split_seconds")
        num_concepts = st.number_input("나눈 구간에서 몇 개의 중요 개념을 뽑을지 입력", value=5, key="num_concepts")

    # URL 입력을 저장할 리스트
    url_list = []

    # 10개의 URL 입력 칸 생성
    for i in range(10):
        url_input = st.text_input(f"URL {i+1} 입력")
        url_list.append(url_input)

    # 제출 버튼
    if st.button("제출"):
        # 감성분석 및 주요 개념 추출 후 엑셀 파일 생성
        Analysis(st.session_state.split_seconds, st.session_state.num_concepts, url_list)
        
        st.write("분석 결과 엑셀 파일 다운로드")
        # 엑셀 파일 다운로드 버튼 생성
        with open('combined_excel.xlsx', 'rb') as f:
            st.download_button('Download result_excel', f, file_name='result_excel.xlsx')
        
        print("===========분석이 끝났습니다.===========")

# 테스트 URL
# https://www.youtube.com/watch?v=5dPK_OGU9FY
# https://www.youtube.com/watch?v=PAvL_NRH8QQ
# https://www.youtube.com/watch?v=blcebnMf2Oc
