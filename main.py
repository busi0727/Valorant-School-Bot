import discord
import datetime
import random
from discord import app_commands
import discord, asyncio
from discord.ext import commands, tasks
from discord import Interaction
from discord.ui import Button, View
from discord.ext import commands
from discord.utils import get
from discord import Intents
from aiohttp import request
import os
from openpyxl import load_workbook, Workbook
import time
from apscheduler.schedulers.background import BackgroundScheduler

token = 'token = open("token", "r").readline()'

############################게임 확률 명령어###############################

suc = -1
per = 0

def play_game():
    global suc
    per = random.randrange(1, 10001)
    if 0 < per <= 5000:
        suc = 0
        print(f'실패')

    elif 5000 < per <=9000:
        suc = 1
        print(f'싱글킬 성공')

    elif 9000 < per <=9700:
        suc = 2
        print(f'더블킬 성공')

    elif 9700 < per <=9900:
        suc = 3
        print(f'클러치 성공')

    elif 9900 < per <=9990:
        suc = 4
        print(f'팀에이스 성공')

    elif 9990 < per <=9999:
        suc = 5
        print(f'에이스 성공')

    elif 9999 < per <=10000:
        suc = 6
        print(f'퍼펙트 에이스 성공')


###############################점수 추가 명령어######################################

def add_score(_name, row, add_amount):
    loadFile()
    print(str(_name), "의 점수를 추가합니다.")
    ws.cell(row, c_score).value += add_amount
    print(_name, "의 점수 : " + str(ws.cell(row, c_score).value))
    ws.cell(row, c_day_count).value -= 1
    print(_name, "의 횟수 : " + str(ws.cell(row, c_day_count).value))
    print("------------------------------\n")
    saveFile()

#################################메인 명령어################################

class aclient(discord.Client):
    def __init__(self):
        super().__init__(intents = discord.Intents.default())
        self.synced = False

    async def on_ready(self):
        await self.wait_until_ready()
        if not self.synced: 
            await tree.sync() 
            self.synced = True
        print(f'{self.user}이 시작되었습니다')  #  봇이 시작하였을때 터미널에 뜨는 말
        game = discord.Game('Visual Studio Code')          # ~~ 하는중
        await self.change_presence(status=discord.Status.idle, activity=game)


client = aclient()
tree = app_commands.CommandTree(client)
sched = BackgroundScheduler()
wb = load_workbook("userDB.xlsx")
ws = wb.active

############################회원가입 명령어#################################

c_name = 1
c_id = 2
c_score = 3
c_day_count = 4

def checkRow():
    for row in range(2, ws.max_row + 1):
        if ws.cell(row,1).value is None:
            return row
            break
def checkName(_name, _id):
    for row in range(2, ws.max_row+1):
        if ws.cell(row,1).value == _name and ws.cell(row,2).value == _id:
            break
            return False
        else:
            return True
            break

    saveFile()

def loadFile():
    wb = load_workbook("userDB.xlsx")
    ws = wb.active

def saveFile():
    wb.save("userDB.xlsx")
    wb.close()

def checkFirstRow():
    print("checkFirstRow")
    loadFile()

    print("첫번째 빈 곳을 탐색")

    for row in range(2, ws.max_row + 1):
        if ws.cell(row,1).value is None:
            return row
            break

    _result = ws.max_row+1

    saveFile()

    return _result

def checkUserNum():
    print("checkUserNum")
    loadFile()

    count = 0

    for row in range(2, ws.max_row+1):
        if ws.cell(row,c_name).value != None:
            count = count+1
        else:
            count = count
    return count

def checkUser(_name, _id):
    print("checkUser")
    print(str(_name) + "<" + str(_id) + ">의 존재 여부 확인")
    print("")

    loadFile()

    userNum = checkUserNum()
    print("등록된 유저수: ", userNum)
    print("")

    print("이름과 고유번호 탐색")
    print("")

    for row in range(2, 3+userNum):
        print(row, "번째 줄 name: ", ws.cell(row,c_name).value)
        print("입력된 name: ", _name)
        print("이름과 일치 여부: ", ws.cell(row, c_name).value == _name)

        print(row,"번째 줄 id: ", ws.cell(row,c_id).value)
        print("입력된 id: ", hex(_id))
        print("고유번호정보와 일치 여부: ", ws.cell(row, c_id).value == hex(_id))
        print("")

        if ws.cell(row, c_name).value == _name and ws.cell(row,c_id).value == hex(_id):
            print("등록된  이름과 고유번호를 발견")
            print("등록된  값의 위치: ",  row, "번째 줄")
            print("")

            saveFile()

            return True, row
            break
        else:
            print("등록된 정보를 탐색 실패, 재탐색 실시")

    saveFile()
    print("발견 실패")

    return False, None

def Signup(_name, _id):
    print("signup")

    loadFile()

    _row = checkFirstRow()
    print("첫번째 빈곳: ", _row)
    print("")

    print("데이터 추가 시작")

    ws.cell(row=_row, column=c_name, value=_name)
    print("이름 추가 | ",  ws.cell(_row,c_name).value)
    ws.cell(row=_row, column=c_id, value =hex(_id))
    print("고유번호 추가 | ", ws.cell(_row,c_id).value)
    ws.cell(row=_row, column=c_score, value= 0)
    print("점수 추가 | ", ws.cell(_row,c_score).value)
    ws.cell(row=_row, column=c_day_count, value= 1)
    print("일일 횟수 추가 | ", ws.cell(_row,c_day_count).value)

    print("")

    saveFile()

    print("데이터 추가 완료")

@tree.command(name = '회원가입', description='게임 참여를 위해 회원가입을 진행합니다.') 
async def login(interaction: discord.Interaction):
    userExistance, userRow = checkUser(interaction.user.name, interaction.user.id)
    if userExistance:
        print("DB에서 ", interaction.user.name, "을 찾았습니다.")
        print("------------------------------\n")
        await interaction.response.send_message(f"이미 회원가입이 되어있습니다.") 
    else:
        print("DB에서 ", interaction.user.name, "을 찾을 수 없습니다")
        print("")

        Signup(interaction.user.name, interaction.user.id)

        print("회원가입이 완료되었습니다.")
        print("------------------------------\n")
        await interaction.response.send_message(f"회원가입이 완료되었습니다.")

#####################################메인 게임 메시지 명령어#################################

@tree.command(name = '게임시작', description='게임을 시작합니다.') 
async def Game_start(interaction: discord.Interaction):
    userExistance, userRow = checkUser(interaction.user.name, interaction.user.id)
    button1 = Button(label="시작하기", style = discord.ButtonStyle.green)
    async def button_callback1(interaction:discord.Interaction):
        play_game()
        if userExistance:
            if ws.cell(userRow, c_day_count).value >= 1:
                await interaction.response.send_message("발사합니다!")
                if suc == 0:
                    view.clear_items()
                    await interaction.channel.send(embed = discord.Embed(title='실패..', 
                    description="적이 도망갔습니다.",colour=discord.Colour.red()), 
                    view=view)

                elif suc == 1:
                    view.clear_items()
                    add_score(interaction.user.name, userRow, 1)
                    await interaction.channel.send(embed = discord.Embed(title='싱글킬!', 
                    description="짧게 피킹 하는 적을 처치했습니다.",colour=discord.Colour.blue()), 
                    view=view)

                elif suc == 2:
                    view.clear_items()
                    add_score(interaction.user.name, userRow, 2)
                    await interaction.channel.send(embed = discord.Embed(title='더블킬!', 
                    description="A롱에서 들어오는 적 두명을 처치했습니다.",colour=discord.Colour.blue()), 
                    view=view)

                elif suc == 3:
                    view.clear_items()
                    add_score(interaction.user.name, userRow, 5)
                    await interaction.channel.send(embed = discord.Embed(title='클러치!', 
                    description="상대방을 처치하고 스파이크를 해체했습니다.",colour=discord.Colour.blue()), 
                    view=view)

                elif suc == 4:
                    view.clear_items()
                    add_score(interaction.user.name, userRow, 10)
                    await interaction.channel.send(embed = discord.Embed(title='팀 에이스!', 
                    description="팀원이 모두 한명씩 처치했습니다",colour=discord.Colour.blue()), 
                    view=view)

                elif suc == 5:
                    view.clear_items()
                    add_score(interaction.user.name, userRow, 30)
                    await interaction.channel.send(embed = discord.Embed(title='에이스!', 
                    description="다시하기.",colour=discord.Colour.blue()), 
                    view=view)

                elif suc == 6:
                    view.clear_items()
                    add_score(interaction.user.name, userRow, 100)
                    await interaction.channel.send(embed = discord.Embed(title='퍼펙트 에이스!', 
                    description="다시하기.",colour=discord.Colour.blue()), 
                    view=view)
            else:
                await interaction.response.send_message("오늘 횟수를 다 사용하셨습니다. 내일 다시 시도해 주세요.")
        else:
            await interaction.response.send_message("회원가입후 이용해주세요.")

    '''
        -성공 명령어
        (embed = discord.Embed(title='성공!', 
        description="다시하기.",
        colour=discord.Colour.blue()), 
        view=view)

        -실패 명령어
        (embed = discord.Embed(title='실패..', 
        description="다시하기.",
        colour=discord.Colour.red()), 
        view=view)
    '''
    button1.callback = button_callback1
    view = View()
    view.add_item(button1)
    await interaction.response.send_message(
        embed = discord.Embed(title='오퍼레이터 사격 게임', 
        description=f"남은 횟수 : {ws.cell(userRow, c_day_count).value} \n 오퍼레이터로 사격하여 적을 죽이는 게임입니다",
        colour=discord.Colour.dark_gray()), 
        view=view
        )

##################################도움(Help) 명령어################################

@tree.command(name="도움말", description="봇을 사용하기 위한 도움말을 보여줘요.")
async def help(interaction: discord.Interaction):
    guild = client.get_guild(1053299673206640711)
    gi = guild.icon
    embed=discord.Embed(title=f'<a:_heart:1061559734081171476> 안녕하세요? <a:_heart:1061559734081171476>',description="저는 발로란트 학교를 위한 다목적 봇이에요!", color=43690)
    embed.add_field(name="/도움말 <a:star_1:1061559907595329568>",value=f"`봇을 사용하기 위한 도움말을 보여줘요.`",inline=False)
    embed.add_field(name="/회원가입 <a:star_2:1061559966835691590>",value=f"`오퍼레이터 사격 게임을 하기 위한 회원가입 명령어에요.`",inline=False)
    embed.add_field(name="/게임시작 <a:star_3:1061559984955080714>",value=f"`오퍼레이터 사격 게임을 플레이하는 명령어에요.`",inline=False)
    embed.add_field(name="/내정보 <a:star_1:1061559907595329568>",value=f"`닉네임, 점수 등 본인의 정보를 확인할 수 있는 명령어에요.\n현재 개발중이에요!`:wrench:",inline=False)
    embed.set_thumbnail(url=f"{gi}")
    embed.set_author(name="Valorant School Bot", icon_url=f"{gi}")
    embed.set_footer(text=f"버그 발생 시, 부시#0150로 문의해주세요.\nMade with ♥ By 부시#0150(536026459471609906)")
    await interaction.response.send_message(embed=embed)

#####################################시간 체크 테스트##################################

def day_count_reset(all_user):
    loadFile()
    for row1 in range(2, 2 + int(all_user)):
        ws.cell(row1, c_day_count).value = 1
        print(row1, "번의 day_count 초기화 완료")
    print("초기화가 끝났습니다.")
    print("------------------------------\n")
    saveFile()

def check_daytime():
    now = datetime.datetime.now()
    loadFile()
    userNum1 = checkUserNum()
    print("등록된 유저수: ", userNum1)
    print(" ")
    print ("날짜 : ", now.strftime("%Y / %m / %d"))
    print(" ")
    print("하루가 지났습니다.")
    print("day_count를 초기화합니다.")
    print("------------------------------\n")
    day_count_reset(int(userNum1))
    saveFile()

@sched.scheduled_job('interval', seconds=5, id='test_1')
def job1():
    print(f'Now Time : {time.strftime("%H:%M:%S")}')

@sched.scheduled_job('cron', hour='15', minute='50', id='test_2')
def hour_check():
    print(f'Now Time : {time.strftime("%H:%M:%S")}, Initiallizing Game_play_count...')
    check_daytime()

@tree.command(name = '테스트', description='안뇽!?!?!?!!!?!?!???!?!?!?') 
async def test(interaction: discord.Interaction):
    await interaction.response.send_message("debuging...")
    check_daytime()

sched.start()

###################################### 내정보 확인 하기 ####################################

@tree.command(name="내정보", description="닉네임, 점수 등 본인의 정보를 확인할 수 있는 명령어에요.")
async def my_Information(interaction: discord.Interaction):
    userExistance, userRow = checkUser(interaction.user.name, interaction.user.id)
    if not userExistance:
        print(f"DB에서 {interaction.user.name}을 찾지 못했습니다.")
        print("------------------------------\n")
        await interaction.response.send_message(f"회원가입 후 이용해주세요.") 
    else:
        guild = client.get_guild(1053299673206640711)
        gi = guild.icon
        iun = interaction.user.name
        iua = interaction.user.avatar
        embed=discord.Embed(title=f'<a:check_8:1060790808388841522> {interaction.user.name}님의 프로필 <a:check_8:1060790808388841522>',description="ㅤ", color=43690)
        embed.add_field(name="<:star1:1061560017746137119> 닉네임",value=f"{interaction.user.name}",inline=False)
        embed.add_field(name="<:star2:1061560043914399814> 점수",value=f"{ws.cell(userRow, c_score).value}점",inline=False)
        embed.add_field(name="<:star3:1061560074335694888> 남은 기회",value=f"{ws.cell(userRow, c_day_count).value}번",inline=False)
        embed.add_field(name="<:star4:1061560100311023697> 랭킹",value="미완성",inline=False)
        embed.add_field(name="<:star5:1061560147085905930> 교장쌤 바보",value="인정",inline=False)
        embed.add_field(name="<:star6:1061560169965817916> 교장쌤 멍청이",value="ㄹㅇㅋㅋ",inline=False)
        embed.set_thumbnail(url=f"{gi}")
        embed.set_author(name=f"{iun}", icon_url=f"{iua}")
        embed.set_footer(text=f"버그 발생 시, 부시#0150로 문의해주세요.\nMade with ♥ By 부시#0150(536026459471609906)")
        await interaction.response.send_message(embed=embed)

######################################토큰#######################################

client.run(token)
