#!/usr/bin/env python
# -*- coding: utf-8 -*-
import stanza
import docx2txt
import re
from konlpy.tag import Mecab, Hannanum, Okt
import sys
import re
import pymongo
from pymongo import MongoClient

#mongoDB 가져오기
mongoURI ="mongodb+srv://user:aa123123@cluster0.bv7lq.mongodb.net/myFirstDatabase?retryWrites=true&w=majority"
client = MongoClient(mongoURI)
db = client.get_database('nlp_DB')
db.collection_names()

mecab = Mecab('C:/mecab/dic/mecab-ko-dic') #mecab import

# 리스트 초기화
subject_dict = ["제+가", "저+", "아빠+가", "아버지+가", "아저씨+가", "나+는", "내+가","저+는"]
triple_dicts=[]
new_dicts=[]
negative = ""
anaphora_dicts=[]
preceding_sentence=[]
direction_keyword = ["쪽에", "에", "왼쪽"]
direction_word =""
logical_structure =[]
attacker=["아저씨","아빠","큰아빠","할아버지","고모","삼촌","남자친구"]

#3.세부정보의 풍부함
quantity_of_detail = ["잡다","끌다","잡아당기다","붙잡히다","만지다","벗기다","벗다","넣다","쑤시다","푸르다", "다가오다","뒤따르다","때리다","들어오다","눕히다","끌어안다","들이대다","키스","눕다","막다","잤다"]
quantity_of_detail_verb_original = ["인사했는데", "구경시켜주고","배고프냐고", "먹겠다고","라면 끓여주고", "앉아있었어요","모른 척 했는데", "물어봐서", "알려줄 수 있냐고", "알려달라고 했는데", "알려줘가지고", "누웠는데", "얘기하고", "잤어요","막았어요","문 열고", "앉았어요","먹고 있었어요","적어줬어요","같이 있었어요","계속 있었는데","누우래요", "누웠는데", "가서", "쉬었어요","움직였어요", "누워서", "안고", "모르고", "해주겠다고","누워있다가", "하지말라고", "아프다고", "뿌리치고", "나왔거든요", "그랬었어요", "비틀고", "피했는데", "뛰쳐나오려고", "끌어들이는", "미안하다고", "잡아서", "깔려", "깼어요",  "달려가려고","소리치면서", "울었는데", "자는 척", "끝났어요","뒤집고", "일어나서", "뒤척였어요", "뿌리치고", "소리를 지르고", "나갔거든요", "찼어요",  "마셨어요", "취하면", "취했어요", "비추는 거예요", "비치잖아요", "빛나고", "쭈구리고","접어서", "거예요", "취해서","틀어막다",  "숨기다", "이야기하지" ,"하더라고요","숙이고", "주저앉", "엎드려", "널브러지듯이", "던져갖고",  "입었어요", "끝났어요", "사주고", "좋아하거든요"]
quantity_of_detail_attacker=["만져서", "만졌어요", "안고", "뽀뽀했어요","만지려고 했어요","만지는 거예요","댔어요","핥았어요", "대서", "대가지고", "넣어가지고", "끌어안으려고", "터치하고", "넣으려고", "벗기는", "더듬고", "싸겠다고", "빨고", "주무르다","주무르듯이","들키다", "집어넣어서", "때리다","벌려","세게","찔렀어요","간지럽혔어요", "올렸어요", "빨아먹다", "쓰다듬듯이","죽여", "죽일까","머리채","협박","싸대기"]
quantity_of_detail_feeling=["이상했어요", "딱딱", "아픈", "싫고", "수치스럽고", "기분", "더럽고","세게", "놀라서","피곤", "당황스럽고", "싫잖아요", "편한", "예민", "부끄러운", "표정","무서웠어요", "느낌"]
quantity_of_detail_body=["쭈쭈","입", "혀", "귀","몸","목","뺨", "얼굴" ,"머리","머리카락", "입술","혀", "볼","혓바닥","뒤통수","이마","손","손등","손바닥","양손","손가락","왼손","오른손","손목", "팔꿈치", "팔","팔목", "왼팔", "오른팔", "주먹","어깨","허벅지","다리", "종아리", "무릎", "발목","등", "골반", "허리","배","배꼽", "겨드랑이"]
quantity_of_detail_person=["씨", "동생", "언니" ,"오빠", "아빠", "엄마", "그 사람", "형", "남자친구", "시아버지","아들","친구","부모님", "작은 애", "큰 애", "고모부", "할머니","고모","이모","누나","남편", "아저씨", "사람", "삼촌","숙모", "아버지", "어머니", "아줌마","선생님","코치","부인","할아버지","신랑","어르신","쌤"]
quantity_of_detail_relationship=["관계", "사이", "둘", "남남", "셋", "세 명", "두 명"]
quantity_of_detail_looknsurround=["옷차림", "잠바", "바지", "반팔", "잠옷", "털 있는", "부드럽고", "옷", "속옷","이불", "얇다", "긴팔", "나시", "두꺼운", "두껍진", "맨투맨", "상의", "청바지","맨몸", "베개", "패딩", "쫄티", "반바지", "면청", "티", "헐렁", "스판", "런닝", "모자", "야구복", "면티","치마", "스타킹", "조끼", "난방", "스포츠", "지퍼","문신", "긴바지", "핑크색","검은색","노란색","티셔츠","양말","윗도리", "7부","내복", "말투","TV", "텔레비전", "탁자","거울", "휴지", "포대", "핸드폰","폰", "휴대폰","전화","통화", "톡","의자","트렁크","근처","모텔","산속","이어폰","테이프","휠체어","티슈"]
quantity_of_detail_surroundnbackground=["마을","논","논밭","침대", "안방", "쇼파", "방바닥", "맨바닥", "땅바닥","바닥", "마루", "거실", "잠결", "창문", "부엌", "방", "침대",  "불이 켜져", "상태", "공간","아파트","옥상","계단", "편의점", "빛","벽","소리", "운전","터미널", "집", "상황","화장실", "방앗간", "앞문", "뒷문", "대문","쓰레기통","물티슈", "엘리베이터","연기", "소파", "병원","공원","슈퍼","테이블","베란다", "살던 곳", "옛날", "자전거","사는 곳","냉장고", "주방","스케이트","신발", "놀이터", "마트","과자","가게","스탠드", "책상", "정수기", "간호실","시장","은행","정류장","식당","부대", "대학교","버스","사회복지관", "건물","짜장면집","노래방","조수석","운전석","공부방", "고기 집","신발장","카운터","당구장","백화점", "인형","동영상","컴퓨터"]
quantity_of_detail_timenbackground=["날씨", "추웠", "춥다", "해가 있을 때", "있을 때", "않을 때", "때", "때쯤","주말",  "엄마가 없었던 날", "새해","아침", "시간","해지기 전", "오전", "오후", "후","그 날", "다음 날","저녁", "밤", "아침", "낮", "점심", "새벽", "초저녁", "더웠는지","아무도","층","하루","이틀","늦은","이른","어두웠어요","어둡다", "밝다", "깜깜","1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월","1일","2일","3일", "4일", "5일", "6일", "7일", "8일", "9일", "10일","11일","12일", "13일", "14일", "15일", "16일", "17일", "18일", "19일", "20일", "21일","22일","23일", "24일", "25일", "26일","27일","28일","29일","30일","31일","1시", "2시","3시","4시","5시","6시","7시","8시","9시","10시","11시","12시","30초", "5초", "10초","3초","30분", "시간" ,"2분", "40분","10분", "5분","3분","잠깐","짧게", "길게","20분","월요일", "화요일", "수요일", "목요일","금요일", "토요일", "일요일","일주일","개월","년","년도","처음으로","도중에", "따뜻", "춥지", "중간에", "다음에","있다가", "이따가","천천히", "빠르게",]
quantity_of_detail_count=["한 번","두 번","세 번","다섯 번", "네 번","열 번", "여러 번", "몇 번", "1개","2개","3개","4개","5개"]
quantity_of_detail_position=["오른쪽", "왼쪽", "일자로", "정자세", "아래", "속에","뒤에","뒤로","자세","위에","위로", "위를", "올려가지고", "밑을", "밑으로", "밑쪽", "쪽", "배쪽으로", "쪽으로", "쪽에서", "앞으로", "앞에", "올라와서", "들어와서", "내려와서", "안으로", "밖에", "옆으로", "똑바로", "정면으로", "여기에", "안쪽", "양쪽", "옆에", "속으로", "바깥쪽","깊숙이", "얕게", "깊숙한", "향하게"]
quantity_of_detail_school=["중 3", "중3", "중2", "중 2", "중1", "중 1","고 1", "고1", "고2","고 2", "고3", "고 3", "초등학교", "학년","방학","시험", "중학교", "고등학교","기말","교복", "미성년자","1학기", "2학기","고사"]
quantity_of_detail_region=["강릉", "부천", "여수", "강남","부산","거창","남천동","함양","공주","신림","서울","인천","순천",]
quantity_of_detail_mainkeyword=["생식기", "아기", "낳는", "섹스", "성교육", "가슴","팬티", "질", "엉덩이", "골","성기", "브래지어", "주물", "추행","키스","반항", "야동","19세", "더듬", "쭈물","콘돔", "젖꼭지","쪼물딱", "꼭지", "삽입", "조몰락", "유륜", "유두", "부위", "팬티", "뽀뽀", "협박", "성관계", "브라자", "브라","입맞춤", "사정", "조물딱", "쪽쪽", "정액","음순" ,"고추","성폭행", "성추행","야한", "고추","중요부위", "소중한 부분","쉬","음부", "소중한 곳", "소중", "성폭력", "말랑","잠지","꼬추","거시기","자위", "터치","모텔","성 경험","강간","사타구니","마사지","꼬치","성경험"]
quantity_of_detail_subkeyword=["원샷","발버둥", "반복","아둥바둥","억지로", "강제","만지작", "필름", "종이컵", "소주" ,"맥주", "시늉","술", "이렇게", "생리", "끄덕", "피", "저번처럼","이렇게", "임신", "살살", "소변", "허겁지겁", "오줌", "구멍", "티격태격", "내동댕이", "액","털","돈", "만 원", "천 원", "0원", "진동", "강요", "담배","캔", "몽롱", "헐레벌떡", "어리버리", "정색", "똑같이", "막막", "행동","흔들", "꿈틀","고함","어깨동무","스토킹","욕"]

#4.맥락상 깊이(공간적 맥락, 시간적 맥락, 연대기적 맥락, 행동 맥락, 심리적 맥락)
depth_in_context_timekeyword = ["주말","아침","점심","저녁","밤","새벽","오전","오후", "봄","여름","가을","겨울","월요일","화요일","수요일","목요일","금요일","토요일","일요일","1시","2시","3시","4시","5시","6시","7시","8시","9시","10시","11시","12시"]
depth_in_context_timekeyword2=["어두컴컴하다","어둡다","밝다"]
depth_in_context_locationkeyword =["방","방바닥","맨바닥", "거실","집","침대","초등학교","방바닥","아파트","옥상","계단","학교","마을","버스터미널","식탁"]
depth_in_context=["느껴지다","이상하다","무섭다","춥다"]


p = re.compile('(\d{4})년|(\d{2})년|([1-9])월|(1[012])월|([1-9])일|([12][0-9])일|(3[01])일|([1-9])시|(1[012])시|\d{1,2}:\d{1,2}|초([1-6])|중([1-3])|고([1-3])|초등학교([1-6])학년|중학교([1-3])학년|고등학교([1-3])학년')
#5. 상호작용의 묘사
attacker_nonverbal=["잡다","끌다","잡아당기다","붙잡히다","만지다","벗기다","벗다","넣다","쑤시다","푸르다", "다가오다","뒤따르다","삽입","찌르다","때리다","관두다","들어오다","눕히다","꽂다","끌어안다","들이대다","멈추다","치다","힘주다","대다","키스"]
attacker_verbal=["벌리라고","가만있으라고"]
attacker_nonverbal_original=["하더라고요","거예요","계셨어요","하셔가지고"]
attacker_link_word=["키스 하다","입 막다","술 먹이다","입 틀어막다"]
victim_nonverbal=["뿌리치다","도망치다", "피하다", "벗겨지다","밀다","소리치다","밀치다", "막다","잠그다","오므리","아프다","울다","잘라내다","도망가다","달려가다","나가다","나오다","업히다","치우다","움직이다","뛰다","발버둥"]
victim_verbal =["하지 말라고","아프다고", "그만하라고","싫다고","안 한다","저리 가라고","그러라고"]
victim_nonverbal_original =["가만히 있었는데", "이렇게 있었는데","업혔는데","치웠는데","손 쳤고","손만 쳤죠","끌려들어","애를 썼어요","탁 쳤거든요"]
victim_link_word=["손 치다","도망 치다","도망","소리 치다","뒷걸음질 치다","손 치우다","고함 지르다","티 내다","가만히 있다","발버둥 치다","소리 지르다","소리 내다"]
#6.대화의 재현
reproduction_of_conversation =["는 거예요","하라고","아프다고","한다고","물었","물어","말라고", "그러면서"] #대화 내용의 구체적이고 연속적인 묘사
#12.주관적인 감정묘사
emotion_dict=["생각","느낌","느껴지다","이상하다", "당황하다","당황", "불편하다", "불쾌하다", "불안하다", "수치","놀라다", "놀래다","부끄럽다", "겁나다", "무섭다", "속상하다", "답답하다", "더럽다", "싫다", "열받다", "짜증나다", "아프다", "어리둥절하다","나쁘다","참다","간지럽다"]
emotion=[]
#15.기억부족의 시인["기억이 안 나는데","기억이 안 나서","기억이 안나","잘 기억이 안나","기억 안 나는데","기억이 안 나고","기억은 안나","기억 안 나고"]
admitting_lack_of_memory=["기억","안"]
#16.자기 진술에 대한 의심 제기: 자신이 한 진술이 진실이 아니거나 부정확할 수 있다는 걱정, 의심
raising_doubts_about_own_testimony=["잘 모르겠","아마","아닐 수도"]

#staza download
print("Downloading Korean model...")
stanza.download('ko', package="gsd")
#stanza pipeline download
print("Building an Korean pipeline...")
ko_nlp = stanza.Pipeline('ko',  package="gsd")

#파일 가져오기(파일 이름 변경 ex. CBCA_신빙성있음_<파일번호>)
text=docx2txt.process('C:/Users/lab/Desktop/stanza/documents/CBCA_신빙성있음_3.docx')

raw_sentence=text.replace('\n\n','\n')
raw_sentence=raw_sentence.replace('\t<','<')
raw_sentence=raw_sentence.replace('\t답\t','답\t')
raw_sentence=raw_sentence.replace('\t문\t','문\t')
raw_sentence=raw_sentence.replace('\t답 ','답\t')
raw_sentence=raw_sentence.replace('\t문 ','문\t ')
raw_sentence=raw_sentence.replace('\t분석관\t','분석관\t')
raw_sentence=raw_sentence.replace('\t피해자\t','피해자\t')

#녹취록 전문을 확인하고 싶을때는 주석 제거
#print(raw_sentence)

def SetExtraction(text):
  new_text = text.replace('[', '')
  text=new_text
  new_text = text.replace(']', '')
  text=new_text
  pattern = r'\{[^}]*\}'
  new_text = re.sub(pattern=pattern, repl='', string= text)
  new_text = new_text.splitlines()
  Set=[]
  for i in range(len(new_text)-1):
    if len(new_text[i])>0 and len(new_text[i+1])>0:
      if new_text[i][0]=='문' and new_text[i+1][0]=='답':
        if new_text[i][1]=='\t' and new_text[i][2]!='\t':
          if new_text[i][-1]=='\t':
            new_text[i]=new_text[i][:-1]
          if new_text[i+1][-1]=='\t':
            new_text[i+1]=new_text[i+1][:-1]
          Set.append([new_text[i][2:],new_text[i+1][2:]])
          i+=1
      if new_text[i][0]=='분' and new_text[i+1][0]=='피':
        if new_text[i][3]=='\t' and new_text[i][4]!='\t':
          if new_text[i][-1]=='\t':
            new_text[i]=new_text[i][:-1]
          if new_text[i+1][-1]=='\t':
            new_text[i+1]=new_text[i+1][:-1]
          Set.append([new_text[i][4:],new_text[i+1][4:]])
          i+=1
      if new_text[i][0]=='분' and new_text[i+1][0]=='피':
        if new_text[i][3]=='\t' and new_text[i][4]=='\t':
          if new_text[i][-1]=='\t':
            new_text[i]=new_text[i][:-1]
          if new_text[i+1][-1]=='\t':
            new_text[i+1]=new_text[i+1][:-1]
          Set.append([new_text[i][5:],new_text[i+1][5:]])
          i+=1
      if new_text[i][0]=="<":
        Set.append([new_text[i][0:],new_text[i][0:]])
  return Set

s=SetExtraction(raw_sentence)


def lemmatize_Han(phrase):
  morphtags = Hannanum().pos(phrase)
  morphtags = [(m +'다' if t.startswith('P') else m + '다') for m, t in morphtags]
  return morphtags

def convert_positive(text):
  # print('질문:',text)
  # print('Mecab', mecab.pos(text))

  rev_sent=list((mecab.pos(text)))

  sent=[]

  for items in rev_sent:
    items=list(items)
    sent.append(items)
  # print(sent)

  result_list=[]
  test = ''

  for items[::-1] in sent:
        # 예외 동사
        # 예외 동사
      if items[1] =='NNB' and items[0] =='거':
        result= text[::-1].replace(items[0][::-1],'것이다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
      elif items[1] =='VV+EP+EC+VX' and items[0] =='했었잖':
        result= text[::-1].replace(items[0][::-1],'했다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
        # 예외 동사
      elif items[1] =='EC+VX+EF' and items[0] =='잖아요':
        result= text[::-1].replace(items[0][::-1],'다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
      elif items[1]=='VV+EF' and items[0]==('기억나'):
        result=text[::-1].replace(items[0][::-1],'기억나다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1]=='VX' and items[0]==('가지'):
        result= text[::-1].replace(items[0][::-1],'가지고이다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
      elif items[1]=='VA+EF' or items[1]=='VV+EF':
        result=text[::-1].replace(items[0][::-1],lemmatize_Han(items[0])[0][::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1] == 'VCP+EF' and items[0] =='인가요':
        result=text[::-1].replace(items[0][::-1],'이다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1] == 'VCP+EF' and items[0] == '인가':
        result=text[::-1].replace(items[0][::-1],'이다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1] == 'VCP+EF' or items[1] == 'EC' or items[1] == 'JX' :
        result=text[::-1].replace(items[0][::-1],'다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1] =='EC+VX':
        result= text[::-1].replace(items[0][::-1],'다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
      elif items[1] =='EF' and items[0] != '냐고' and items[0] != '다고':
        result=text[::-1].replace(items[0][::-1],'다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1] == 'JKO':
        result=text[::-1].replace(items[0][::-1],'이다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1] =='EF' and items[0] == '냐고' or items[0] == '다고':
        result=text.strip('?')
        result='\''+result+'\''+'이다.'
        break;
      elif items[1]== 'JKS' or items[1] == 'NNG' or items[1] == 'VV+ETM' or items[1] == 'NP' or items[1] == 'JKB' or items[1] == 'JKS' or items[1] == 'NNBC' or items[1] == 'NNB' or items[1] =='MAG' or items[1] =='XSN' or items[1] =='EC' or items[1] == 'JX' or items[1] == 'JKO' or items[1] == 'VV+EC':
        result=text.strip('?')
        result=result+'이다.'
        break;
      elif items[1]=='VX+EF' or items[1]=='XSA+EF':
        result=text[::-1].replace(items[0][::-1],'하다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1]=='NNB+VCP+EF' and items[0].startswith('건'):
        result=text[::-1].replace(items[0][::-1],'것이다'[::-1],1)
        result=result[::-1].replace('?','.')
        break;
      elif items[1]=='JX' and items[0]==('요'):
        result=text[::-1].replace(items[0][::-1],'다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
      else: result = '질문에 대한 긍정'

  return result

def convert_negative(text):
  # print('질문:',text)
  global result

  rev_sent=list((mecab.pos(text)))

  sent=[]

  for items in rev_sent:
    items=list(items)
    sent.append(items)
  # print(sent)

  result_list=[]
  test = ''

  for items[::-1] in sent:
    if items[0]=='상태':
        result= text[::-1].replace(items[0][::-1],'상태가 아니다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='NNB' and items[0] =='거':
        result= text[::-1].replace(items[0][::-1],'것이 아니다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='NNB' and items[0] =='게':
        result= text[::-1].replace(items[0][::-1],'것이 아니다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='VV' and items[0] =='알':
        result= text[::-1].replace(items[0][::-1],'모른다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1]== 'NNBC' or items[1]== 'JKB' or items[1]== 'MAG':
        result=text.strip('?')
        result=result+'이(가) 아니다.'
        break;
    elif items[1]=='VV+EP':
        result=text[::-1].replace(items[0][::-1],lemmatize_Han(items[0])[0][::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='VA' and items[0] =='있':
        result= text[::-1].replace(items[0][::-1],'없다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='VCP+EF' and items[0] == '예요':
        result= text[::-1].replace(items[0][::-1],'가 아니다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='VCP+EF' and items[0] == '나요':
        result= text[::-1].replace(items[0][::-1],'나지 않는다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='EF'and items[0] != '냐고' and items[0] != '다고':
        result= text[::-1].replace(items[0][::-1],'지 않다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='EF' and items[0] == '냐고' or items[0] == '다고':
      result=text.strip('?')
      result='\''+result+'\''+'가 아니다.'
      break;
    elif items[1] =='EC' and items[0].startswith('던가'):
        result= text[::-1].replace(items[0][::-1],'지 않다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='EC' and items[0].startswith('거나'):
        result= text[::-1].replace(items[0][::-1],'지 않다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1] =='ETN+JX' and items[0] == '긴':
        result= text[::-1].replace(items[0][::-1],'지 않았다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
        break;
    elif items[1]=='VX+EF' and items[0] =='줄래요':
        result= text[::-1].replace(items[0][::-1],'주기 싫다.'[::-1],1)
        result=result[::-1].split(".")[0]+"."
    else: result='질문에 대한 부정'
  return result

def convert(s):
  for i in s:
    # print(i[1])
    if i[1].startswith("네")>0 or i[1].startswith("예")>0 or i[1].startswith("(고개를 끄덕이다)")>0:
      i[1]=convert_positive(i[0])
      if i[1]=='질문에 대한 긍정':
        i[1]='"'+i[0]+'"라는 '+i[1]
    elif i[1].startswith("아니")>0 or i[1].startswith("(고개를 좌우로 흔들다)")>0 or i[1].startswith("(고개를 젓는다)")>0 or i[1].startswith("(고개를 저으")>0:
      i[1]=convert_negative(i[0])
      if i[1]=='질문에 대한 부정':
        i[1]='"'+i[0]+'"라는 '+i[1]

#의문문 평서문 변환
convert(s)

#문답에서 '문' question 리스트에 저장
question=[]
for i in s:
  question.append(i[0])

#문답에서 '답' question 리스트에 저장
result=[]
for i in s:
  result.append(i[1])


change_result = ""
for i in result:
  change_result += i
# stanza 문장 분석
ko_doc = ko_nlp(change_result)
#  형태소 분석기 okt
okt = Okt()

# Triple 추출
count = 0
for i, sent in enumerate(ko_doc.sentences):
    j = i
    print("[Sentence {}]".format(i + 1))

    triple_dict = {'주어': 'X', '목적어': 'X', '서술어': 'X', "관형어": "X", "부사어": "X", "방향": "X"} #triple dict 초기화
    original_dict = {'주어': 'X', '목적어': 'X', '서술어': 'X'}
    print()

    for i, word in enumerate(sent.words):
        if word.text in direction_keyword or "쪽" in word.text or "에" in word.text and "나중" not in word.text and "처음" not in word.text:
            if sent.words[i - 1].deprel == "case" and sent.words[i - 1].pos == "NOUN" or (
                    sent.words[i - 1].deprel == "nmod" and sent.words[i - 1].pos == "NOUN") or (
                    sent.words[i - 1].deprel == "advmod" and sent.words[i - 1].pos == "NOUN"):
                direction_word = sent.words[i - 1].text + " " + word.text
            if direction_word != "":
                print(direction_word)
                triple_dict['방향'] = direction_word
                direction_word = ""
            else:
                direction_word = word.text
                print(direction_word)
                triple_dict['방향'] = direction_word
                direction_word = ""

        if (word.deprel == "nsubj" or word.deprel == "csubj"):
            if (word.xpos == "NP+JKS"):
                print("{}번째 목적어 : {:15s}".format(j + 1, word.text))
                continue
            if (word.xpos == "NNB"):
                print("{}번째 의존명사 : {:15s}".format(j + 1, word.text))
                continue
            if (word.xpos == "NNB+JKC"):
                print("{}번째 의존명사 : {:15s}".format(j + 1, word.text))
                continue
            if (word.lemma in subject_dict):
                print("{}번째 주요 주어 : {:15s}".format(j + 1, word.text))
                triple_dict['주어'] = word.text
                continue
            print("{}번째 주어 : {:15s}".format(j + 1, word.text))
        #  triple_dict['주어'] = word.text
        elif (word.deprel == "det"):
            print("{}번째 관형사 : {:15s}".format(j + 1, word.text))
            triple_dict['관형어'] = word.text

        elif (word.deprel == "dep"):
            if (word.pos == "VERB"):
                print("{}번째 서술어 : {:15s}".format(j + 1, word.text))
                triple_dict['서술어'] = word.text
                if (triple_dict['주어'] != 'X' and triple_dict['목적어'] != 'X' and triple_dict['서술어'] != 'X'):
                    triple_dicts.append(triple_dict.copy())
                continue
            print("{}번째 목적어 : {:15s}".format(j + 1, word.text))
            triple_dict['목적어'] = word.text
        elif (word.deprel == "acl:relcl"):
            print("{}번째 목적어 : {:15s}".format(j + 1, word.text))
            triple_dict['목적어'] = word.text
        elif (word.deprel == "flat"):
            if sent.words[i - 1].deprel == "advcl" and word.pos == "VERB":
                a = sent.words[i - 1].text + " " + word.text
                print("{}번째 서술어 : {:15s}".format(j + 1, a))
                triple_dict['서술어'] = sent.words[i - 1].text + " " + word.text
                continue
            if (word.pos == "VERB"):
                print("{}번째 서술어 : {:15s}".format(j + 1, word.text))
                triple_dict['서술어'] = word.text
                if negative != "":
                    triple_dict['서술어'] = negative + word.text
                    negative = ""
                else:
                    triple_dict['서술어'] = word.text
                if (triple_dict['주어'] != 'X' and triple_dict['목적어'] != 'X' and triple_dict['서술어'] != 'X'):
                    triple_dicts.append(triple_dict.copy())
                j = j + 1
                continue
            print("{}번째 목적어 : {:15s}".format(j + 1, word.text))
            triple_dict['목적어'] = word.text
        elif (word.deprel == "obj" or word.deprel == "iobj"):
            print("{}번째 목적어 : {:15s}".format(j + 1, word.text))
            triple_dict['목적어'] = word.text
        elif (word.deprel == "ccomp"):
            print("{}번째 서술어 : {:15s}".format(j + 1, word.text))
            triple_dict['서술어'] = word.text
            if negative != "":
                triple_dict['서술어'] = negative + word.text
                negative = ""
            else:
                triple_dict['서술어'] = word.text
            if (triple_dict['주어'] != 'x' and triple_dict['목적어'] != 'X' and triple_dict['서술어'] != 'X'):
                triple_dicts.append(triple_dict.copy())
        elif (word.deprel == "obl"):
            #  if(word.xpos =="NNG+JKB"):
            #    continue
            print("{}번째 관형어 : {:15s}".format(j + 1, word.text))
            triple_dict['관형어'] = word.text

        elif (word.deprel == "conj"):
            print("{}번째 서술어 : {:15s}".format(j + 1, word.text))
            triple_dict['서술어'] = word.text
            if negative != "":
                triple_dict['서술어'] = negative + word.text
                negative = ""
            else:
                triple_dict['서술어'] = word.text
            if (triple_dict['주어'] != 'x' and triple_dict['목적어'] != 'X' and triple_dict['서술어'] != 'X'):
                triple_dicts.append(triple_dict.copy())
            j = j + 1
        #  이거 나중에 고쳐야됨!!
        elif (word.deprel == "advmod" and word.xpos == "MAG" and word.text == "안"):
            print("{}번째 부정지정사 : {:15s}".format(j + 1, word.text))
            negative = word.text
        if (word.deprel == "advcl"):
            if word.pos == "VERB":
                triple_dict['서술어'] = word.text
                if negative != "":
                    triple_dict['서술어'] = negative + word.text
                    negative = ""
                continue
            if ("VCN" in word.xpos):
                print("{}번째 부정지정사 + 서술어 : {:15s}".format(j + 1, word.text))
                triple_dict['목적어'] = 'X'
                j += 1
                continue
            print("{}번째 부사절 : {:15s}".format(j + 1, word.text))
            if triple_dict['부사어'] != "X":
                triple_dict['부사어'] = triple_dict['부사어'] + " " + word.text
            else:
                triple_dict['부사어'] = word.text
            j = j + 1
        elif (word.deprel == "advmod"):
            print("{}번째 부사절 : {:15s}".format(j + 1, word.text))
            #  triple_dict['부사어'] = word.text
            if triple_dict['부사어'] != "X":
                triple_dict['부사어'] = triple_dict['부사어'] + " " + word.text
            else:
                triple_dict['부사어'] = word.text
            j = j + 1
        elif (word.deprel == "root"):
            if (word.pos == "NOUN"):
                print("{}번째 목적어 : {:15s}".format(j + 1, word.text))
                continue
            if word.pos == "ADJ":
                print("{}번째 부사절 : {:15s}".format(j + 1, word.text))
                triple_dict['부사어'] = word.text
                continue
            if ("VCN" in word.xpos):
                print("{}번째 부정지정사 + 서술어 : {:15s}".format(j + 1, word.text))
                triple_dict['목적어'] = 'X'
                continue
            print("{}번째 서술어 : {:15s}".format(j + 1, word.text))
            if word.text == "거예요":
                triple_dict['서술어'] = sent.words[i - 1].text + " " + word.text
            elif negative != "":
                triple_dict['서술어'] = negative + word.text
                negative = ""
            else:
                triple_dict['서술어'] = word.text
            if (triple_dict['주어'] != 'X' and triple_dict['목적어'] != 'X' and triple_dict['서술어'] != 'X'):
                triple_dicts.append(triple_dict.copy())
            j = j + 1
        #  print(triple_dict)
        #  db.victim.insert_one(triple_dict)
        #  triple_dict = {'주어': 'X', '목적어': 'X', '서술어':'X', "관형어":"X", "부사어":"X", "방향" : "X"}

    if triple_dict != {'주어': 'X', '목적어': 'X', '서술어': 'X', "관형어": "X", "부사어": "X", "방향": "X"}:
        db.victim.insert_one(triple_dict)

    if triple_dict["주어"] == 'X' or triple_dict['서술어'] == "X":
        #  print(triple_dict)
        print(sent.text)
    print("")

    try:
        for i in range(len(triple_dicts) - 1):
            if triple_dicts[i]['주어'] == triple_dicts[i + 1]['주어'] and triple_dicts[i]['목적어'] == triple_dicts[i + 1]['목적어']:
                a = triple_dicts[i]['서술어']
                b = triple_dicts[i + 1]['서술어']
                triple_dicts[i]['서술어'] = a + " " + b
                if triple_dicts[i]['관형어'] == "X":
                    triple_dicts[i]['관형어'] = triple_dicts[i + 1]['관형어']
                del triple_dicts[i + 1]
    except:
        pass
    for v in triple_dicts:
        if v not in new_dicts:
            new_dicts.append(v)
            logical_structure.append(v)

db.attacker.drop()

#논리적 일관성 판단 부분 
subject_keyword = ["저","제가","나는","제", "저는"]
sbj_attack = ["아빠가", "아빠"]


for i in db.victim.find():
  sbj = str(i["주어"])
  if sbj in sbj_attack:
    triple_set = {'주어': 'X', '목적어': 'X', '서술어':'X', "관형어":"X", "부사어":"X", "방향" : "X"}
    triple_set['주어'] = sbj
    a = i['서술어']
    triple_set['방향'] = i['방향']
    original=okt.morphs(a, stem= True)
    triple_set['주어'] =i['주어']
    triple_set['목적어'] =i['목적어']
    if original[0] =="안":
      triple_set['서술어'] = original[0] + original[1]
    else:
      triple_set['서술어'] = original[0]

    triple_set['관형어'] =i['관형어']
    triple_set['부사어'] =i['부사어']
    triple_set['방향'] =i['방향']
    db.attacker.insert_one(triple_set)
    print(triple_set)

attacker_db =db.attacker.find()

# 주어와 서술어가 같을 때 방향이 다르면 '논리적 불일치' 출력
for i, word in enumerate(db.attacker.find()):
    for j in range(i):
      try:
        if attacker_db[j]['서술어'] == attacker_db[i]['서술어'] and  attacker_db[j]['서술어'] !="하다" and (attacker_db[i]['방향'] != "X" and attacker_db[j]['방향'] != "X") :
        #  print(word)
         if attacker_db[j]['방향'] != word['방향']:
           print(word)
           print(attacker_db[j])
           print("논리적 불일치")
      except:
        continue
