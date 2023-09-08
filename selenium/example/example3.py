futureCheckList = ["종목정보", "30 종가단일가","종목마감", "M2 당일 확정","정산가격", "현물정보결제기준채권",]

for index in futureCheckList:
    if "종목정보" in index:
        print("O ->>>> 맞음")
    else:
        print("틀림")