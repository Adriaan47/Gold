Dim PureGoldPurity
Dim CurrentGoldPriceInUSD
Dim ExchangeRateZar
Dim GoldItemPurity
Dim GoldItemWeight
Dim TotalGoldInItem
Dim BuyingPrice
Dim ItemPurity
Dim GoldPicePerGram
Dim TroyOunces
Dim OurRate
Dim OurFinalRate
Dim FinalPrice
TroyOunces = 31.1034768
PureGoldPurity = 24
CurrentGoldPriceInUSD = 1800.01 
ExchangeRateZar = 15.71
OurRate = 0.85
OurFinalRate = 0.95


Call CalcGoldBuyingPrice()
Function CalcGoldBuyingPrice()
            'CurrentGoldPriceInUSD=InputBox("Enter today's gold price in USD:")
            'ExchangeRateZar=InputBox("Enter today's USD exchange rate")
            GoldItemPurity=InputBox("How many carats is the gold item?","Gold Carats /24")
            GoldItemWeight=InputBox("What does the gold items weight?","Weight is grams") 
            GoldPicePerGram =(CurrentGoldPriceInUSD / TroyOunces) 
            ItemPurity = (GoldItemPurity / PureGoldPurity)
            TotalGoldInItem = (GoldItemWeight * ItemPurity)
            BuyingPrice = (TotalGoldInItem * GoldPicePerGram * ExchangeRateZar)
            MsgBox("The amount of gold in the item is: " & TotalGoldInItem & " grams") 
            FinalPrice = (BuyingPrice * OurRate)
            
            MsgBox("We will purchase the gold for: R" & Round(FinalPrice,2))
            MsgBox("Does the customer Agree?")
            MsgBox("Final offer")
            MsgBox("We will purchase the gold for: R" & (BuyingPrice * OurFinalRate))
            MsgBox("We can sell for")
            MsgBox (BuyingPrice)
End Function







