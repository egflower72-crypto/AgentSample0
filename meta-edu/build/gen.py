import sys, json, os, re, hashlib
sys.path.insert(0, os.path.dirname(__file__))
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CACHE = os.path.join(HERE, "emoji_cache"); os.makedirs(CACHE, exist_ok=True)

# ---- 연락처/링크 (한 곳에서 관리) ----
CONTACT_NAME = "호잔"
CONTACT_TEL = "010-3818-1713"
CONTACT_MAIL = "link2plus1@gmail.com"
SITE = "link21premium.com"
FOOT = f"호잔 {CONTACT_TEL}  ·  {CONTACT_MAIL}  ·  {SITE}"

BLUE=RGBColor(0x00,0x66,0xFF); GREEN=RGBColor(0x22,0xC5,0x5E); DKG=RGBColor(0x10,0xB9,0x81)
ORANGE=RGBColor(0xF3,0x72,0x2C); NAVY=RGBColor(0x26,0x46,0x53); CHAR=RGBColor(0x2C,0x3E,0x50)
GRAY=RGBColor(0x71,0x80,0x96); LIGHT=RGBColor(0xE8,0xF4,0xF8); WHITE=RGBColor(255,255,255)
BG=RGBColor(0xF6,0xFA,0xFD); PALEG=RGBColor(0xE9,0xFB,0xEF); PALEO=RGBColor(0xFF,0xF0,0xE6)
FONT="Noto Sans KR"

EMOJI_FONT = ImageFont.truetype("/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf", 109)
def split_emoji(s):
    # 이모지 시퀀스(ZWJ, variation selector 포함)를 낱개로 분리
    out=[]; cur=""
    for ch in s:
        if ch in "‍️" : cur+=ch; continue
        if cur and not cur.endswith("‍"):
            out.append(cur); cur=""
        cur+=ch
    if cur: out.append(cur)
    return out
def emoji_png(e):
    h=hashlib.md5(e.encode()).hexdigest()[:12]
    p=os.path.join(CACHE,h+".png")
    if not os.path.exists(p):
        im=Image.new("RGBA",(160,140),(0,0,0,0)); d=ImageDraw.Draw(im)
        d.text((0,0),e,font=EMOJI_FONT,embedded_color=True)
        im=im.crop(im.getbbox() or (0,0,160,140)); 
        # 정사각형 패딩
        w,hh=im.size; s=max(w,hh); sq=Image.new("RGBA",(s,s),(0,0,0,0)); sq.paste(im,((s-w)//2,(s-hh)//2)); sq.save(p)
    return p

prs = Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
BLANK = prs.slide_layouts[6]

def rect(sl,x,y,w,h,fill,line=None,shape=MSO_SHAPE.ROUNDED_RECTANGLE,rad=0.08,lw=1.0):
    s=sl.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h))
    if shape==MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0]=rad
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb=fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb=line; s.line.width=Pt(lw)
    s.shadow.inherit=False
    return s
def text(sl,x,y,w,h,t,size=16,bold=False,color=CHAR,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,name=None):
    tb=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)); tf=tb.text_frame
    tf.word_wrap=True; tf.margin_left=tf.margin_right=Inches(0.05); tf.margin_top=tf.margin_bottom=Inches(0.03)
    tf.vertical_anchor=anchor
    lines=t if isinstance(t,list) else [t]
    for i,l in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align
        runs=l if isinstance(l,list) else [(l,{})]
        for rt,ro in runs:
            r=p.add_run(); r.text=rt; f=r.font; f.name=FONT; f.size=Pt(ro.get("size",size)); f.bold=ro.get("bold",bold)
            f.color.rgb=ro.get("color",color)
    if name: tb.name=name
    return tb
def emo(sl,x,y,size,e,center_w=None):
    parts=split_emoji(e)
    n=len(parts); gap=size*0.12
    tot=n*size+(n-1)*gap
    x0=x+((center_w-tot)/2 if center_w else 0)
    for i,pt in enumerate(parts):
        pic=sl.shapes.add_picture(emoji_png(pt),Inches(x0+i*(size+gap)),Inches(y),Inches(size),Inches(size))
        pic.name="emoji "+pt
def bg(sl):
    sl.background.fill.solid(); sl.background.fill.fore_color.rgb=BG
def logo(sl,x=0.4,y=0.22,w=1.1,h=0.42):
    r=rect(sl,x,y,w,h,GREEN,rad=0.3); 
    tf=r.text_frame; tf.margin_left=tf.margin_right=0; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    ru=p.add_run(); ru.text="LINK21"; ru.font.name="Inter"; ru.font.size=Pt(14); ru.font.bold=True; ru.font.color.rgb=WHITE
    tf.vertical_anchor=MSO_ANCHOR.MIDDLE
def header(sl,area,tag,page,total):
    bg(sl); logo(sl)
    text(sl,1.7,0.2,8.5,0.45,area,14,True,NAVY,anchor=MSO_ANCHOR.MIDDLE)
    t=rect(sl,10.5,0.22,2.4,0.42,BLUE,rad=0.3)
    tf=t.text_frame; tf.vertical_anchor=MSO_ANCHOR.MIDDLE; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    ru=p.add_run(); ru.text=tag; ru.font.name=FONT; ru.font.size=Pt(13); ru.font.bold=True; ru.font.color.rgb=WHITE
    footer(sl,page,total)
def footer(sl,page,total):
    text(sl,0.4,7.08,9.5,0.3,FOOT,11,False,GRAY)
    text(sl,11.4,7.08,1.5,0.3,f"{page}/{total}" if total else str(page),11,False,GRAY,PP_ALIGN.RIGHT)
def notes(sl,t): sl.notes_slide.notes_text_frame.text=t

def ans_letters(a): return a.replace("  /  "," · ")

def flow_row(sl,flow,y,h=2.4,x0=0.5,w_total=12.3):
    n=len(flow); arrow_w=0.55
    cw=(w_total-(n-1)*arrow_w)/n
    for i,(e,lab) in enumerate(flow):
        x=x0+i*(cw+arrow_w)
        rect(sl,x,y,cw,h,WHITE,LIGHT,lw=1.5)
        ne=len(split_emoji(e)); esz=min(1.0,(cw-0.4)/ne-0.1)
        emo(sl,x,y+0.2,esz,e,center_w=cw)
        text(sl,x+0.1,y+0.2+esz+0.1,cw-0.2,h-esz-0.4,lab.split('\n'),16 if n<=3 else 14,True,CHAR,PP_ALIGN.CENTER,MSO_ANCHOR.TOP)
        if i<n-1:
            a=rect(sl,x+cw+0.08,y+h/2-0.22,arrow_w-0.16,0.44,BLUE,shape=MSO_SHAPE.RIGHT_ARROW)

def cover(total):
    sl=prs.slides.add_slide(BLANK); bg(sl)
    rect(sl,0,0,13.333,7.5,BLUE,shape=MSO_SHAPE.RECTANGLE)
    rect(sl,0,5.2,13.333,2.3,RGBColor(0x00,0x4C,0xCC),shape=MSO_SHAPE.RECTANGLE)
    logo(sl,0.7,0.6,1.8,0.7)
    text(sl,0.7,1.7,9,0.5,"소상공인을 위한 · 강사용 교안 (영역 1 견본)",20,False,RGBColor(0xDD,0xEA,0xFF))
    text(sl,0.7,2.25,11,1.3,"Meta 인증 모의고사\n쉽게 풀어주는 사장님 교실",44,True,WHITE)
    text(sl,0.7,4.1,9,0.5,"영역 1 · Meta 기술의 가치 (시험 비중 8%) · 32문항",20,True,WHITE)
    emo(sl,9.3,1.6,1.1,"📘"); emo(sl,10.6,1.6,1.1,"📸"); emo(sl,11.9,1.6,1.0,"💬")
    text(sl,0.7,5.55,11.5,0.5,"강사 연락처",14,False,RGBColor(0xDD,0xEA,0xFF))
    text(sl,0.7,5.95,12,0.6,f"{CONTACT_NAME}  ·  {CONTACT_TEL}  ·  {CONTACT_MAIL}",24,True,WHITE)
    text(sl,0.7,6.6,12,0.5,f"Meta 과정 오픈 시 신청 안내  →  {SITE}",18,False,WHITE)
    notes(sl,f"표지. 강사 소개 후 '오늘은 어려운 용어를 사장님 말로 바꿔서 풀어봅니다'로 시작. 연락처: {CONTACT_NAME} {CONTACT_TEL} / {CONTACT_MAIL}")

def guide(total,pg):
    sl=prs.slides.add_slide(BLANK); header(sl,"교안 사용법","강사 가이드",pg,total)
    text(sl,0.5,0.9,12,0.8,"이 교안은 이렇게 쓰세요",30,True,BLUE)
    rows=[("00~10분","도입","사장님 가게 고민 이야기 (문제 슬라이드의 말풍선 읽기)"),
          ("10~30분","이론","핵심 개념 3장: Meta 앱 가족도 · 발견+연결 · 한 번에 여러 앱"),
          ("30~60분","실습 1","문제 슬라이드 → 손들기 → 풀이 슬라이드로 정답 확인"),
          ("60~90분","실습 2","틀린 문제만 다시. 사장님 가게 사례로 바꿔 말하기"),
          ("90~120분","마무리","체크리스트 + 과정 신청 안내")]
    for i,(t,n,d) in enumerate(rows):
        y=1.8+i*0.82
        rect(sl,0.5,y,12.3,0.72,WHITE,LIGHT)
        text(sl,0.7,y,1.8,0.72,t,16,True,BLUE,anchor=MSO_ANCHOR.MIDDLE)
        text(sl,2.5,y,1.5,0.72,n,18,True,CHAR,anchor=MSO_ANCHOR.MIDDLE)
        text(sl,4.0,y,8.7,0.72,d,15,False,CHAR,anchor=MSO_ANCHOR.MIDDLE)
    rect(sl,0.5,6.05,12.3,0.85,PALEO,ORANGE)
    text(sl,0.7,6.05,12,0.85,[[("강사 주의 ",{"bold":True,"color":ORANGE}),("소상공인 2시간 실습 중심(이론 2 : 실습 8). 성과를 보장한다는 말, 과장 광고, 고객 개인정보 입력은 수업 중 계속 짚어 주세요.",{})]],14,anchor=MSO_ANCHOR.MIDDLE)
    notes(sl,"링크21 강의자료 규칙(소상공인): 이론 2 : 실습 8, 친근체, 무료 도구 중심, 과장 광고 위험 명시. 문제당 2장(문제→풀이). 노트에 강사 멘트와 그림 설명법이 있습니다.")

def area_cover(total,pg):
    sl=prs.slides.add_slide(BLANK); header(sl,"영역 1 · Meta 기술의 가치","영역 표지",pg,total)
    text(sl,0.5,0.9,12,0.9,"영역 1 · Meta 기술의 가치",34,True,BLUE)
    text(sl,0.5,1.75,12,0.5,"시험 비중 8% · 32문항 · 쉬운 영역이니 여기서 자신감을 챙기세요",16,False,GRAY)
    text(sl,0.5,2.4,6,0.5,"오늘의 학습 목표",20,True,NAVY)
    goals=[("🔎","Meta 앱 6가지가 각각 무슨 일을 하는지 말할 수 있다"),("🏪","우리 가게에 맞는 앱을 고를 수 있다"),("📝","'아닌 것 고르기'·'알맞은 기능 고르기' 문제를 풀 수 있다")]
    for i,(e,g) in enumerate(goals):
        x=0.5+i*4.15
        rect(sl,x,2.95,3.95,3.7,WHITE,LIGHT,lw=1.5)
        emo(sl,x,3.2,1.2,e,center_w=3.95)
        text(sl,x+0.2,4.6,3.55,1.9,f"{i+1}. {g}",18,True,CHAR,PP_ALIGN.CENTER)
    notes(sl,"학습목표 3개(3의 법칙). 측정 가능한 동사: 말할 수 있다/고를 수 있다/풀 수 있다. 각 카드를 짚으며 읽어 주세요.")

def concept_family(total,pg):
    sl=prs.slides.add_slide(BLANK); header(sl,"영역 1 · 핵심 개념 ①","핵심 개념",pg,total)
    text(sl,0.5,0.85,12,0.7,"내 가게 손님은 Meta 어디에 있을까?",30,True,BLUE)
    apps=[("📘","Facebook","그룹·Marketplace·Watch"),("📸","Instagram","사진·영상으로 발견"),("💬","Messenger","1:1 실시간 대화"),
          ("🧾","WhatsApp","암호화·무료 메시지·카탈로그"),("🧵","Threads","텍스트 중심 대화"),("🌐","Audience Network","외부 앱에도 광고")]
    for i,(e,n,d) in enumerate(apps):
        c=i%3; r=i//3; x=0.5+c*4.15; y=1.75+r*2.55
        rect(sl,x,y,3.95,2.3,WHITE,LIGHT,lw=1.5)
        emo(sl,x+0.25,y+0.35,1.2,e)
        text(sl,x+1.6,y+0.4,2.3,0.6,n,20,True,BLUE,anchor=MSO_ANCHOR.MIDDLE)
        text(sl,x+1.6,y+1.0,2.3,1.1,d,15,False,CHAR)
    text(sl,0.5,6.85,12.3,0.3,"+ Meta Quest(VR·MR 헤드셋)도 Meta 기술이에요. YouTube는 Meta가 아니에요!",13,True,ORANGE)
    notes(sl,"[그림 설명법] 6개 카드를 왼쪽 위부터 '사장님 가게에서 이 앱은 어떤 역할?' 하고 한 장씩 질문하며 넘깁니다. 예: Facebook=동네 반상회·장터, Instagram=가게 간판 사진, Messenger/WhatsApp=카운터 상담, Audience Network=이웃 상가에 전단지 붙이기. 시험 포인트: Meta 앱 아닌 것은 YouTube.")

def concept_value(total,pg):
    sl=prs.slides.add_slide(BLANK); header(sl,"영역 1 · 핵심 개념 ②","핵심 개념",pg,total)
    text(sl,0.5,0.85,12,0.7,"Meta가 가게에 주는 두 가지: 발견 + 연결",30,True,BLUE)
    flow_row(sl,[("🔎","발견\n새 손님이 우리 가게를 알게 됨"),("🤝","연결\n손님과 대화하며 관계를 이어감"),("🏪","내 가게 성장")],1.8,h=2.9)
    rect(sl,0.5,5.0,12.3,1.6,WHITE,LIGHT,lw=1.5)
    text(sl,0.8,5.1,11.8,1.4,[[("발견 도구: ",{"bold":True,"color":BLUE}),("Instagram 릴스·탐색 탭, 제품 태그",{})],[("연결 도구: ",{"bold":True,"color":GREEN}),("Messenger·WhatsApp·Instagram DM, 라이브, 방송 채널",{})],[("한곳 관리: ",{"bold":True,"color":NAVY}),("Meta Business Suite",{})]],19,anchor=MSO_ANCHOR.MIDDLE)
    notes(sl,"[그림 설명법] 가운데 화살표를 짚으며 '발견→연결→성장' 순서로 읽습니다. 사장님께 '우리 가게는 지금 발견이 약한가, 연결이 약한가?'를 한 명씩 물어보면 참여가 올라갑니다. 시험 문제 27번의 정답이 이 두 단어입니다.")

def concept_multi(total,pg):
    sl=prs.slides.add_slide(BLANK); header(sl,"영역 1 · 핵심 개념 ③","핵심 개념",pg,total)
    text(sl,0.5,0.85,12,0.7,"광고 한 번 설정으로 여러 앱에 동시에",30,True,BLUE)
    flow_row(sl,[("🍜","캠페인 1개"),("📘📸💬🌐","Facebook · Instagram · Messenger · Audience Network"),("📈","넓은 도달")],1.8,h=2.9)
    rect(sl,0.5,5.0,12.3,1.8,PALEO,ORANGE)
    text(sl,0.8,5.1,11.8,1.6,[[("꼭 기억해요 ",{"bold":True,"color":ORANGE})],"• AI 자동화(어드밴티지+)는 시간을 줄여줄 뿐, 성과를 보장하지 않아요.","• Meta Verified(인증 배지)는 광고 심사를 건너뛰게 해주지 않아요."],16)
    notes(sl,"[그림 설명법] '한 번 차려서 여러 상에 낸다'는 말로 설명. 18번·29번·28번 문제의 오답 함정(보장/면제)을 미리 짚어 줍니다.")

def q_slide(q,c,page,total):
    sl=prs.slides.add_slide(BLANK)
    area=q["area"]; no=q["no"]
    header(sl,area,f"문제 {q['id']}/32",page,total)
    # 좌: 가게 장면
    rect(sl,0.5,0.95,4.7,3.55,WHITE,LIGHT,lw=1.5)
    emo(sl,0.5,1.1,1.5,c["shop"][0],center_w=4.7)
    text(sl,0.6,2.65,4.5,0.4,c["shop"][1]+" 사장님",16,True,BLUE,PP_ALIGN.CENTER)
    b=rect(sl,0.75,3.1,4.2,1.25,PALEG,GREEN)
    text(sl,0.85,3.1,4.0,1.25,"“"+c["situ"]+"”",14,False,CHAR,anchor=MSO_ANCHOR.MIDDLE)
    # 우: 쉬운 질문
    text(sl,5.5,0.95,7.3,0.4,"쉽게 풀어 쓴 질문",13,True,GRAY)
    text(sl,5.5,1.3,7.3,1.9,c["easy"],26,True,CHAR)
    rect(sl,5.5,3.3,7.3,1.2,LIGHT,None)
    text(sl,5.6,3.3,7.1,1.2,[[("시험 원문  ",{"bold":True,"color":GRAY,"size":11}),(q["q"],{"size":12,"color":GRAY})]],12,anchor=MSO_ANCHOR.MIDDLE)
    # 보기 2x2
    for i,o in enumerate(q["opts"]):
        col=i%2; row=i//2; x=0.5+col*6.2; y=4.7+row*1.15
        rect(sl,x,y,6.1,1.05,WHITE,BLUE,lw=1.2)
        cc=rect(sl,x+0.15,y+0.25,0.55,0.55,BLUE,shape=MSO_SHAPE.OVAL)
        tf=cc.text_frame; tf.margin_left=tf.margin_right=0; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; ru=p.add_run(); ru.text=o[0]; ru.font.name=FONT; ru.font.size=Pt(16); ru.font.bold=True; ru.font.color.rgb=WHITE
        body=o[3:] if o[1:3]==". " else o
        text(sl,x+0.8,y,5.2,1.05,body,15 if len(body)<45 else 13,False,CHAR,anchor=MSO_ANCHOR.MIDDLE)
    notes(sl,f"[진행] 1) 왼쪽 말풍선을 사장님 목소리로 읽어 줍니다. 2) 오른쪽 '쉬운 질문'을 읽고 3) 20초 정도 생각하게 한 뒤 손들기. 정답은 다음 장에서 공개합니다.\n[시험 원문] {q['q']}")

def a_slide(q,c,page,total):
    sl=prs.slides.add_slide(BLANK)
    header(sl,q["area"],f"풀이 {q['id']}/32",page,total)
    text(sl,0.5,0.8,12.3,0.4,"그림으로 이해하기",14,True,GRAY)
    flow_row(sl,c["flow"],1.2,h=2.5)
    # 정답
    rect(sl,0.5,3.95,6.0,2.95,PALEG,GREEN,lw=2)
    text(sl,0.7,4.0,5.6,0.4,"정답",14,True,DKG)
    text(sl,0.7,4.4,5.6,0.9,c["key"],26 if len(c["key"])<22 else 20,True,CHAR)
    text(sl,0.7,5.55,5.6,1.3,[[("시험 보기  ",{"bold":True,"color":GRAY,"size":11}),(ans_letters(q["ans"]),{"size":13,"color":GRAY})]],13)
    rect(sl,6.7,3.95,6.1,1.45,WHITE,BLUE,lw=1.2)
    text(sl,6.85,4.0,5.8,1.4,[[("💡 쉬운 이유  ",{"bold":True,"color":BLUE})],c["why"]],14)
    rect(sl,6.7,5.5,6.1,1.4,PALEO,ORANGE,lw=1.2)
    text(sl,6.85,5.55,5.8,1.3,[[("⚠️ 이런 함정 조심  ",{"bold":True,"color":ORANGE})],c["trap"]],14)
    chain=" → ".join(l.replace("\n"," ") for _,l in c["flow"])
    notes(sl,f"[그림 설명법] 위 그림을 왼쪽부터 짚으며 '{chain}' 순서로 읽어 주세요.\n[강사 멘트] {c['say']}\n[원문 해설] {q['exp']}")

def checklist(total,pg):
    sl=prs.slides.add_slide(BLANK); header(sl,"영역 1 마무리","마무리",pg,total)
    text(sl,0.5,0.85,12,0.7,"오늘 이것만 기억하세요",30,True,BLUE)
    items=["Meta 앱 6가지: Facebook · Instagram · Messenger · WhatsApp · Threads · Audience Network","Meta 앱이 아닌 것: YouTube (Google 서비스)","Meta가 주는 두 가지 가치: 발견 + 연결","한 캠페인으로 여러 앱에 동시 게재 가능","AI 자동화·인증 배지: 성과·심사 면제를 보장하지 않는다"]
    for i,t in enumerate(items):
        y=1.7+i*0.72
        rect(sl,0.5,y,12.3,0.62,WHITE,LIGHT)
        c=rect(sl,0.65,y+0.1,0.42,0.42,GREEN,shape=MSO_SHAPE.OVAL)
        tf=c.text_frame; tf.margin_left=tf.margin_right=0; tf.vertical_anchor=MSO_ANCHOR.MIDDLE; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER; ru=p.add_run(); ru.text="✓"; ru.font.size=Pt(14); ru.font.bold=True; ru.font.color.rgb=WHITE; ru.font.name=FONT
        text(sl,1.25,y,11.4,0.62,t,16,False,CHAR,anchor=MSO_ANCHOR.MIDDLE)
    rect(sl,0.5,5.5,12.3,1.4,PALEO,ORANGE)
    text(sl,0.7,5.5,11.9,1.4,[[("AI·보안 한마디 ",{"bold":True,"color":ORANGE})],"AI(Meta AI 등)가 만든 문구는 사장님이 사실 확인 후 쓰기 · 손님 전화번호·주소 같은 개인정보는 대화·AI 창에 함부로 입력하지 않기 · 과장 광고 금지"],15,anchor=MSO_ANCHOR.MIDDLE)
    notes(sl,"체크리스트를 소리 내어 같이 읽습니다. 마지막에 AI 윤리·보안 한 가지를 꼭 짚습니다(링크21 강의자료 규칙: 윤리 매 수업 포함).")

def cta(total,pg):
    sl=prs.slides.add_slide(BLANK); bg(sl)
    rect(sl,0,0,13.333,7.5,BLUE,shape=MSO_SHAPE.RECTANGLE)
    logo(sl,0.7,0.6,1.8,0.7)
    text(sl,0.7,1.7,12,1.2,"Meta 과정이 열리면\n가장 먼저 신청하세요",42,True,WHITE)
    rect(sl,0.7,3.6,11.9,2.6,WHITE,None)
    text(sl,1.0,3.7,11.3,0.5,"신청·문의",16,True,GRAY)
    text(sl,1.0,4.15,11.3,0.7,f"🌐  {SITE}",30,True,BLUE)
    text(sl,1.0,4.9,11.3,0.6,f"📞  {CONTACT_NAME}  {CONTACT_TEL}",24,True,CHAR)
    text(sl,1.0,5.5,11.3,0.6,f"✉️  {CONTACT_MAIL}",24,True,CHAR)
    text(sl,0.7,6.45,12,0.5,"※ 과정 오픈 일정·신청 페이지는 오픈 시 안내드립니다.",14,False,RGBColor(0xDD,0xEA,0xFF))
    notes(sl,"과정은 아직 오픈 전이므로 '오픈되면 안내'로 말합니다. 정확한 일정·신청 URL이 확정되면 이 장의 문구만 수정하세요(상단 FOOT/SITE 상수는 build/gen.py).")

def build(out):
    d=json.load(open(os.path.join(ROOT,"data/q400.json")))
    import content_a1
    qs={x["id"]:x for x in d if x["id"]<=32}
    cm={x["id"]:x for x in content_a1.C}
    total=2+1+3+ 32*2 +2+1  # 대략치
    pages=[]
    # 먼저 총 슬라이드 수 계산 위해 두 번 돌리지 않고 고정
    total=1+1+1+3+64+2
    cover(total)
    guide(total,2); area_cover(total,3); concept_family(total,4); concept_value(total,5); concept_multi(total,6)
    pg=7
    for i in range(1,33):
        q_slide(qs[i],cm[i],pg,total); pg+=1
        a_slide(qs[i],cm[i],pg,total); pg+=1
    checklist(total,pg); pg+=1; cta(total,pg)
    prs.save(out); print("saved",out,len(prs.slides))

if __name__=="__main__":
    build(os.path.join(ROOT,"out","Meta_영역1_소상공인_강사용교안_견본.pptx"))
