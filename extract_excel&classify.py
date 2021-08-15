import win32com.client
import pandas as pd
import statistics

targetCodeList = []  # 종목군별 소속 종목 코드를 담은 2차원 배열
targetindustry = [132, 133, 144, 145, 146, 147, 148]  # 종목군
# 132 KOSPI200 산업재 ,133 KOSPI200 헬스케어 ,144 KOSPI200 에너지/화학,
# 145 KOSPI200 정보기술,146 KOSPI200 금융,
# 147 KOSPI200 생활소비재, 148 KOSPI200 경기소비재


instCpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
instMarketEye = win32com.client.Dispatch('CpSysDib.MarketEye')

# 엑셀데이터 인덱싱후, 종목별 평균 per(=avg_per) 구하기
xlsx_data = pd.read_excel(r'stock_value.xlsx', header=8)
cols = xlsx_data.columns

for i in range(1, len(cols)):
    if i % 2 == 1:
        del xlsx_data[cols[i]]

xlsx_data = xlsx_data.fillna(0)
xlsx = xlsx_data.iloc[5:, 1:]
xls = pd.DataFrame(xlsx)

mean_df = xls.mean()
# avg_per 종목별 평균 per을 나타내는 데이터프레임
avg_per = pd.DataFrame(mean_df)

for i in targetindustry:
    targetCodeList.append(
        instCpCodeMgr.GetGroupCodeList(i))  # instCpCodeMgr.GetGroupCodeList() : 업종코드에 해당하는 종목 코드 리스트 반환

# per_list 종목별 평균 per을 담은 리스트
per_list = []
for i in range(len(targetCodeList)):
    ind = []
    for j in range(len(targetCodeList[i])):
        code = targetCodeList[i][j]
        ind_list = avg_per[0][code + '.1']
        ind.append(ind_list)
    per_list.append(ind)

average_per_list = []  # 종목군 별 평균 per을 담은 list
for i in range(len(targetindustry)):
    average_per_list.append(statistics.mean(per_list[i]))

down = []  # 종목군 평균 per 이하의 종목을 담은 리스트
up = []  # 종목군 평균 per 이상의 종목을 담은 리스트

industry = input("원하는 종목군을 입력 ( 0 : KOSPI200 산업재 ,1 : KOSPI200 헬스케어 , "\
                 "2: KOSPI200 에너지/화학, 3 KOSPI200 정보기술, 4  KOSPI200 금융,\
 5 KOSPI200 생활소비재, 6 KOSPI200 경기소비재)")
print('\n')


# def classify : 입력받은 종목군코드별로 종목군 평균 per 이상/이하별 집단을 분류하는 함수
def classify(a):
    for b in range(len(targetCodeList[a])):
        if per_list[a][b] < average_per_list[a]:
            down.append(per_list[a][b])
        else:
            up.append(per_list[a][b])


classify(int(industry))
print('종목 평균 per 이하 집단', down, '\n')
print('종목 평균 per 이상 집단', up)


# getCAGR : 성과검증을 위해 연평균 성장률을 구하는 함수
def getCAGR(first, last, years):
    return (last / first) ** (1 / years) - 1
