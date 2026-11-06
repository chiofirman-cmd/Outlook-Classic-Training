#!/usr/bin/env python3
"""Build a five-minute Outlook Classic PST training simulation MP4."""
from pathlib import Path
import subprocess, textwrap, os, sys
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'output'
WORK = ROOT / 'work'
OUT.mkdir(exist_ok=True); WORK.mkdir(exist_ok=True)
W, H = 1920, 1080
FPS = 15

SCENES = [
    (20, 'INTRODUCTION', 'Outlook PST Archive Configuration & Rules',
     'Welcome to this training. In the next five minutes, we will configure a monthly Outlook data archive, set the default data file, create general rules for incoming and sent messages, and verify the result.'),
    (50, 'CREATE A NEW PST FILE', 'Create the monthly archive data file',
     'First, open Outlook Classic. On the File tab, select Account Settings, and then choose Account Settings again. Open the Data Files tab and select Add. Save the new Outlook Data File using the agreed naming convention: Archive underscore month and year, followed by dot P S T. For this example, the file name is Archive underscore one one two six dot P S T. The month and year are represented by two digits each. Confirm the file location and add the data file to Outlook.'),
    (35, 'SET THE DEFAULT DATA FILE', 'Select the PST and choose Set as Default',
     'Next, return to the Data Files tab. Select the newly created Archive PST and choose Set as Default. Confirm the change if Outlook displays a prompt. The selected data file is now the default Outlook data file for applicable local data storage. Please note: this setting does not move a Microsoft 365 or Exchange mailbox to the PST, and it does not change the mail server delivery location.'),
    (65, 'CREATE GENERAL INBOX RULES', 'Apply a general rule to incoming messages',
     'Now create the incoming mail rule. From the Home tab, open Rules, then Manage Rules and Alerts, and select New Rule. Choose a rule that applies to messages you receive. Do not select a sender-specific condition. Configure the action to move the message to the specified folder, then select the Inbox folder inside Archive underscore one one two six. Give the rule a clear functional name, such as Inbox Archive Rule, and finish the wizard. Review the rule summary before applying it. In this demonstration, the rule is general and is not limited to a particular sender.'),
    (60, 'CREATE SENT ITEMS RULES', 'Apply a general rule to sent messages',
     'Next, create a separate rule for sent messages. Open Rules and Alerts and start a new rule. Choose the option that applies the rule to messages I send. Do not add a recipient-specific condition. Configure the action to move a copy or route the applicable sent message to the Sent Items folder inside the Archive PST, according to the available Outlook rule options. Name the rule Sent Items Archive Rule. Review the conditions and actions carefully, then save the rule. The Inbox and Sent Items rules remain separate so each workflow is easy to review and maintain.'),
    (50, 'TESTING AND VERIFICATION', 'Verify the archive folders and rule behavior',
     'It is time to test the configuration. Send a sample message and receive a sample message in the Outlook profile. Check the Rules and Alerts window to confirm that both general rules are enabled. Open the Archive PST in the folder pane. Verify that the Inbox and Sent Items destination folders are present and that the sample messages appear in the expected locations. If a rule does not run, check its conditions, destination folder, and whether Outlook is open. Moving messages to a local PST is generally a client-side action and may require classic Outlook to be running. Organization policies and account types can affect the available rule actions.'),
    (20, 'TRAINING COMPLETE', 'PST archive configured and verified',
     'You have completed the Outlook Classic PST archive configuration walkthrough. Remember to back up PST files regularly and keep them in an approved storage location. Thank you for watching.'),
]

FONT_DIR = '/usr/share/fonts/truetype/dejavu'
def font(size, bold=False):
    p = Path(FONT_DIR) / ('DejaVuSans-Bold.ttf' if bold else 'DejaVuSans.ttf')
    return ImageFont.truetype(str(p), size)

def wrap(draw, text, fnt, max_width):
    words=text.split(); lines=[]; line=''
    for word in words:
        test=(line+' '+word).strip()
        if draw.textbbox((0,0),test,font=fnt)[2] <= max_width: line=test
        else:
            if line: lines.append(line)
            line=word
    if line: lines.append(line)
    return lines

def draw_text_block(draw, xy, text, fnt, fill, max_width, line_gap=12):
    x,y=xy
    for line in wrap(draw,text,fnt,max_width):
        draw.text((x,y),line,font=fnt,fill=fill); y += fnt.size + line_gap
    return y

def make_scene(i, title, subtitle):
    im=Image.new('RGB',(W,H),'#f3f5f8'); d=ImageDraw.Draw(im)
    # Windows title bar and Outlook ribbon
    d.rectangle((0,0,W,48),fill='#202b3c')
    d.text((28,10),'Outlook',font=font(22,True),fill='white')
    d.text((W-430,12),'—   □   ×',font=font(22),fill='#dce4ef')
    d.rectangle((0,48,W,148),fill='#ffffff')
    d.text((28,65),'File    Home    Send / Receive    Folder    View    Help',font=font(25),fill='#27364a')
    d.line((0,147,W,147),fill='#cbd3df',width=2)
    d.text((30,105),'New Email     Delete     Archive     Rules     Move     Reply     Forward',font=font(19),fill='#34445a')
    # left folder pane
    d.rectangle((0,150,360,H),fill='#e8edf4')
    d.text((28,180),'FAVORITES',font=font(18,True),fill='#526174')
    items=['Inbox','Sent Items','Drafts','Deleted Items','Archive_1126.pst','   Inbox','   Sent Items']
    yy=225
    for j,item in enumerate(items):
        if (i in (3,4) and j>=4) or (i==5 and j>=4):
            d.rounded_rectangle((12,yy-7,340,yy+34),radius=7,fill='#cfe2fb')
        d.text((35,yy),('▸ ' if j==4 else '   ' if j>4 else '✉ ')+item,font=font(21,j==4),fill='#26384d')
        yy+=58
    # main content area
    d.rectangle((360,150,W,H),fill='#ffffff')
    d.text((410,180),title,font=font(34,True),fill='#17365d')
    d.text((410,235),subtitle,font=font(22),fill='#617187')
    # Outlook mock content varies by scene
    if i==0:
        d.rounded_rectangle((440,330,1450,760),radius=20,fill='#f4f7fb',outline='#cbd5e1',width=3)
        d.text((510,390),'OUTLOOK CLASSIC TRAINING',font=font(38,True),fill='#17365d')
        d.text((510,470),'PST ARCHIVE CONFIGURATION',font=font(30),fill='#2873b9')
        d.text((510,560),'Create  •  Configure  •  Test',font=font(27),fill='#526174')
    elif i==1:
        d.rounded_rectangle((460,320,1460,820),radius=12,fill='#f8fafc',outline='#9aa9bb',width=3)
        d.text((500,350),'Account Settings',font=font(30,True),fill='#26384d')
        d.text((510,420),'Email    Data Files    RSS Feeds    SharePoint Lists',font=font(21),fill='#526174')
        d.line((490,465,1420,465),fill='#b8c4d3',width=2)
        d.text((520,505),'Name',font=font(20,True),fill='#526174'); d.text((930,505),'Location',font=font(20,True),fill='#526174')
        d.rounded_rectangle((510,545,1400,610),radius=6,fill='#dcecff',outline='#5594d6',width=2)
        d.text((530,565),'Outlook Data File',font=font(21),fill='#17365d'); d.text((930,565),'Archive_1126.pst',font=font(21),fill='#17365d')
        d.rounded_rectangle((1220,700,1400,760),radius=8,fill='#1769aa'); d.text((1260,718),'Add...',font=font(22,True),fill='white')
    elif i==2:
        d.rounded_rectangle((470,330,1420,770),radius=12,fill='#f8fafc',outline='#9aa9bb',width=3)
        d.text((520,370),'Outlook Data Files',font=font(29,True),fill='#26384d')
        d.rounded_rectangle((520,460,1360,535),radius=7,fill='#dcecff',outline='#5594d6',width=2)
        d.text((550,482),'Archive_1126.pst',font=font(24),fill='#17365d')
        d.rounded_rectangle((1120,600,1370,670),radius=8,fill='#1769aa'); d.text((1150,620),'Set as Default',font=font(21,True),fill='white')
    elif i in (3,4):
        d.rounded_rectangle((460,310,1500,850),radius=12,fill='#f8fafc',outline='#9aa9bb',width=3)
        d.text((500,345),'Rules and Alerts',font=font(30,True),fill='#26384d')
        d.text((510,415),'Apply this rule after the message arrives',font=font(23),fill='#17365d')
        d.rounded_rectangle((510,475,1430,545),radius=7,fill='#e7f1fc',outline='#8bb6df',width=2)
        d.text((540,495),'Condition: general message handling (no sender/recipient filter)',font=font(19),fill='#26384d')
        dest='Archive_1126.pst  >  Inbox' if i==3 else 'Archive_1126.pst  >  Sent Items'
        d.rounded_rectangle((510,575,1430,655),radius=7,fill='#e7f1fc',outline='#8bb6df',width=2)
        d.text((540,600),'Action: move to  '+dest,font=font(21),fill='#17365d')
        d.text((520,720),'Rule name: '+('Inbox Archive Rule' if i==3 else 'Sent Items Archive Rule'),font=font(23,True),fill='#2873b9')
    elif i==5:
        d.rounded_rectangle((460,320,1480,810),radius=12,fill='#f8fafc',outline='#9aa9bb',width=3)
        d.text((510,360),'Archive_1126.pst',font=font(29,True),fill='#17365d')
        for k,(name,status) in enumerate([('Inbox','Sample incoming message verified'),('Sent Items','Sample sent message verified')]):
            y=450+k*130
            d.rounded_rectangle((520,y,1400,y+90),radius=8,fill='#eef6ef',outline='#a8c8ad',width=2)
            d.text((555,y+14),name,font=font(25,True),fill='#245b32')
            d.text((555,y+52),status,font=font(18),fill='#426b4a')
            d.text((1320,y+22),'✓',font=font(38,True),fill='#2e8b57')
    else:
        d.rounded_rectangle((470,350,1440,760),radius=18,fill='#f4f7fb',outline='#cbd5e1',width=3)
        d.text((560,430),'CONFIGURATION COMPLETED',font=font(35,True),fill='#245b32')
        d.text((560,520),'PST archive  •  Inbox rules  •  Sent Items rules',font=font(23),fill='#526174')
        d.text((560,620),'Thank you for watching',font=font(30),fill='#17365d')
    # bottom caption bar
    d.rectangle((0,1000,W,H),fill='#15243a')
    d.text((40,1022),f'OUTLOOK CLASSIC  |  PST ARCHIVE TRAINING  |  {i+1:02d}/07',font=font(19,True),fill='#dce8f5')
    return im

def run(cmd):
    print('+', ' '.join(map(str,cmd)))
    subprocess.run(list(map(str,cmd)),check=True)

def main():
    if not shutil_which('ffmpeg') or not shutil_which('espeak-ng'):
        raise SystemExit('ffmpeg and espeak-ng are required. The GitHub workflow installs both.')
    segment_videos=[]; audio_files=[]
    for idx,(duration,title,subtitle,narration) in enumerate(SCENES):
        img=WORK/f'scene_{idx:02d}.png'; wav=WORK/f'voice_{idx:02d}.wav'; mp4=WORK/f'segment_{idx:02d}.mp4'
        make_scene(idx,title,subtitle).save(img)
        run(['espeak-ng','-v','en-us','-s','145','-a','165','-w',wav,narration])
        # still-image motion: gentle zoom for a screen-recording/tutorial feel
        frames=duration*FPS
        run(['ffmpeg','-y','-loop','1','-i',img,'-vf',f"zoompan=z='min(zoom+0.00012,1.035)':d={frames}:s=1920x1080:fps={FPS},format=yuv420p",'-t',duration,'-an','-c:v','libx264','-preset','ultrafast','-crf','24',mp4])
        segment_videos.append(mp4); audio_files.append(wav)
    # concat video segments
    concat=WORK/'videos.txt'
    concat.write_text(''.join(f"file '{p.as_posix()}'\n" for p in segment_videos))
    silent=WORK/'video_only.mp4'
    run(['ffmpeg','-y','-f','concat','-safe','0','-i',concat,'-c','copy',silent])
    # concatenate voice clips, pad each scene to its allotted duration, then mux
    audio_inputs=[]; filters=[]
    for i,((duration,*_),wav) in enumerate(zip(SCENES,audio_files)):
        audio_inputs += ['-i',wav]
        filters.append(f'[{i}:a]apad,atrim=0:{duration},asetpts=PTS-STARTPTS[a{i}]')
    filters.append(''.join(f'[a{i}]' for i in range(len(audio_files)))+f'concat=n={len(audio_files)}:v=0:a=1[outa]')
    audio_out=WORK/'narration.wav'
    run(['ffmpeg','-y',*audio_inputs,'-filter_complex',';'.join(filters),'-map','[outa]','-ar','44100','-ac','2',audio_out])
    final=OUT/'Outlook_Classic_Training.mp4'
    run(['ffmpeg','-y','-i',silent,'-i',audio_out,'-map','0:v:0','-map','1:a:0','-c:v','copy','-c:a','aac','-b:a','128k','-t','300','-movflags','+faststart',final])
    print(f'Created {final} ({final.stat().st_size} bytes)')

def shutil_which(name):
    from shutil import which
    return which(name)

if __name__=='__main__': main()
