Attribute VB_Name = "modGetGuruFocusData"
Option Explicit
Public aGuruFocusItems(1 To 250) As Variant
Function smfGetGuruFocusData(ByVal pTicker As String, _
                             ByVal pItem As Integer, _
                    Optional ByVal pPeriod As String = "C", _
                    Optional ByVal pOffset As Integer = 0, _
                    Optional ByVal pError As Variant = "Error", _
                    Optional ByVal pType As Integer = 0) As Variant
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return a data from GuruFocus
   '-----------------------------------------------------------------------------------------------------------*
   ' 2014.12.21 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2017.05.03 -- Change "http://" protocol to "https://"
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample of use:
   '
   '    =smfGetGuruFocusData("MMM",1,"Y1")
   '
   '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    vError = pError
    
    '------------------> Create high-level labels for GuruFocus URL
    Dim sTicker As String, sCompName As String, sCompName2 As String, sURL As String
    Dim sPeriod2 As String
    sTicker = UCase(pTicker)
    Select Case Left(UCase(pPeriod), 1)
       Case "Y": sPeriod2 = " Annual data"
       Case "Q": sPeriod2 = " Quarterly data"
       Case "C": sPeriod2 = "": pOffset = 0
       Case Else: vError = "Improper period -- first byte should be A or Q or C": GoTo ErrorExit
       End Select
    sCompName = smfStrExtr(smfGetTagContent("https://www.gurufocus.com/stock/" & sTicker, "title", 1), "~", " (")
    sCompName2 = Replace(sCompName, " ", "+")
       
    '------------------> Array of potential formulas
    Dim aLabel As Variant
    If aGuruFocusItems(1) = "" Then smfLoadGuruFocusItems
    aLabel = Split(aGuruFocusItems(pItem), "|")
       
    '------------------> Determine which data items to return
    Select Case aLabel(0)
       Case "0": smfGetGuruFocusData = sCompName: Exit Function
       Case "1": GoTo GF_Variation1
       Case Else: vError = "Invalid item number": GoTo ErrorExit
       End Select
    
    Exit Function
    
GF_Variation1:
    sURL = "https://www.gurufocus.com/term/" & aLabel(1) & "/" & sTicker & "/" & aLabel(2) & "/" & sCompName2
    If pItem = 36 Or pItem = 37 Or pItem = 45 Or pItem = 101 Then sPeriod2 = sCompName & " Historical Data"
    Select Case pOffset
       Case 0
            Select Case aLabel(4)
               Case 0: smfGetGuruFocusData = "N/A"
               Case 1: smfGetGuruFocusData = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "div", -1, "As of "), "~", " ("))
               Case 2: smfGetGuruFocusData = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "font", -1, " Rank:"), ": ", "-"))
               End Select
       Case -10 To -1
            Select Case aLabel(3)
               Case "N/A": smfGetGuruFocusData = "N/A"
               Case Else: smfGetGuruFocusData = RCHGetTableCell(sURL, pOffset, sPeriod2, "<tr")
               End Select
       Case 1 To 10
            Select Case aLabel(3)
               Case "N/A": smfGetGuruFocusData = "N/A"
               Case Else: smfGetGuruFocusData = RCHGetTableCell(sURL, -pOffset, sPeriod2, ">" & aLabel(3))
               End Select
       Case Else: vError = "Invalid offset period -- must be -10 to 10": GoTo ErrorExit
       End Select
    Exit Function

ErrorExit: smfGetGuruFocusData = vError
                   
   End Function

Sub smfLoadGuruFocusItems()
       aGuruFocusItems(1) = "0"
       aGuruFocusItems(2) = "1|zscore|Altman%2BZ-Score|X1|0"
       aGuruFocusItems(3) = "1|zscore|Altman%2BZ-Score|X2|0"
       aGuruFocusItems(4) = "1|zscore|Altman%2BZ-Score|X3|0"
       aGuruFocusItems(5) = "1|zscore|Altman%2BZ-Score|X4|0"
       aGuruFocusItems(6) = "1|zscore|Altman%2BZ-Score|X5|0"
       aGuruFocusItems(7) = "1|zscore|Altman%2BZ-Score|Z-Score|1"
       aGuruFocusItems(8) = "1|mscore|Beneish%2BM-Score|DSRI|0"
       aGuruFocusItems(9) = "1|mscore|Beneish%2BM-Score|GMI|0"
       aGuruFocusItems(10) = "1|mscore|Beneish%2BM-Score|AQI|0"
       aGuruFocusItems(11) = "1|mscore|Beneish%2BM-Score|SGI|0"
       aGuruFocusItems(12) = "1|mscore|Beneish%2BM-Score|DEPI|0"
       aGuruFocusItems(13) = "1|mscore|Beneish%2BM-Score|SGAI|0"
       aGuruFocusItems(14) = "1|mscore|Beneish%2BM-Score|LVGI|0"
       aGuruFocusItems(15) = "1|mscore|Beneish%2BM-Score|TATA|0"
       aGuruFocusItems(16) = "1|mscore|Beneish%2BM-Score|M-score|1"
       aGuruFocusItems(17) = "1|fscore|Piotroski%2BF-Score|Q1|0"
       aGuruFocusItems(18) = "1|fscore|Piotroski%2BF-Score|Q2|0"
       aGuruFocusItems(19) = "1|fscore|Piotroski%2BF-Score|Q3|0"
       aGuruFocusItems(20) = "1|fscore|Piotroski%2BF-Score|Q4|0"
       aGuruFocusItems(21) = "1|fscore|Piotroski%2BF-Score|Q5|0"
       aGuruFocusItems(22) = "1|fscore|Piotroski%2BF-Score|Q6|0"
       aGuruFocusItems(23) = "1|fscore|Piotroski%2BF-Score|Q7|0"
       aGuruFocusItems(24) = "1|fscore|Piotroski%2BF-Score|Q8|0"
       aGuruFocusItems(25) = "1|fscore|Piotroski%2BF-Score|Q9|0"
       aGuruFocusItems(26) = "1|fscore|Piotroski%2BF-Score|F-score|1"
       aGuruFocusItems(27) = "1|Book Value Per Share|Book%2BValue%2Bper%2BShare|Book Value Per Share|1"
       aGuruFocusItems(28) = "1|Dividends Per Share|Dividends%2BPer%2BShare|Dividends Per Share|1"
       aGuruFocusItems(29) = "1|per share eps|Earnings%2Bper%2BShare|per share eps|1"
       aGuruFocusItems(30) = "1|eps_nri|Earnings%2Bper%2Bshare%2Bwithout%2BNon-Recurring%2BItems|eps_nri|1"
       aGuruFocusItems(31) = "1|EBITDA_per_share|EBITDA%2Bper%2BShare|EBITDA_per_share|1"
       aGuruFocusItems(32) = "1|ev|Enterprise%2BValue|ev|1"
       aGuruFocusItems(33) = "1|rank_balancesheet|Financial%2BStrength|N/A|1"
       aGuruFocusItems(34) = "1|per share_freecashflow|Free%2BCashflow%2Bper%2BShare|per share_freecashflow|1"
       aGuruFocusItems(35) = "1|e10|E10|N/A|1"
       aGuruFocusItems(36) = "1|InsiderOwnership|Insider%2BOwnership|Insider Ownership|1"
       aGuruFocusItems(37) = "1|InstitutionalOwnership|Institutional%2BOwnership|Institutional Ownership|1"
       aGuruFocusItems(38) = "1|mktcap|Market%2BCap|mktcap|1"
       aGuruFocusItems(39) = "1|netcash|Net%2BCash|netcash|1"
       aGuruFocusItems(40) = "1|NCAV_real|Net%2BCurrent%2BAsset%2BValue%2B%2528per%2Bshare%2529|NCAV|1"
       aGuruFocusItems(41) = "1|NCAV|Net-Net%2BWorking%2BCapital|NNWC|1"
       aGuruFocusItems(42) = "1|predictability_rank|predictability_rank|N/A|2"
       aGuruFocusItems(43) = "1|rank_profitability|Profitability%2BRank|N/A|1"
       aGuruFocusItems(44) = "1|per share rev|Revenue%2Bper%2BShare|per share rev|1"
       aGuruFocusItems(45) = "1|ShortInterest|Short%2BInterest|Short Interest|1"
       aGuruFocusItems(46) = "1|Tangibles_book_per_share|Tangible%2BBook%2BValue%2Bper%2BShare|Tangibles_book_per_share|1"
       aGuruFocusItems(47) = "1|turnover|Asset%2BTurnover|turnover|1"
       aGuruFocusItems(48) = "1|total_buyback_3y|Buyback%2BRatio|N/A|1"
       aGuruFocusItems(49) = "1|CCC|Cash%2BConversion%2BCycle%2B%2528CCC%2529|CCC|1"
       aGuruFocusItems(50) = "1|cash2debt|Cash%2Bto%2BDebt|cash2debt|1"
       aGuruFocusItems(51) = "1|cogs2rev|COGS%2Bto%2BRevenue|cogs2rev|1"
       aGuruFocusItems(52) = "1|current_ratio|Current%2BRatio|current ratio|1"
       aGuruFocusItems(53) = "1|DaysInventory|Days%2BInventory|DaysInventory|1"
       aGuruFocusItems(54) = "1|DaysPayable|Days%2BPayable|DaysPayable|1"
       aGuruFocusItems(55) = "1|DaysSalesOutstanding|Days%2BSales%2BOutstanding|DaysSalesOutstanding|1"
       aGuruFocusItems(56) = "1|deb2equity|Debt%2Bto%2BEquity%2B%2528%2525%2529|deb2equity|1"
       aGuruFocusItems(57) = "1|Earnings Yield|Earnings%2BYield|N/A|1"
       aGuruFocusItems(58) = "1|earning_yield_greenblatt|Earnings%2BYield%2B%2528Joel%2BGreenblatt%2529|earning_yield_greenblatt|1"
       aGuruFocusItems(59) = "1|equity2asset|Equity%2Bto%2BAsset|equity2asset|1"
       aGuruFocusItems(60) = "1|ev2ebit|EV%252FEBIT|ev2ebit|1"
       aGuruFocusItems(61) = "1|ev2ebitda|EV%252FEBITDA|ev2ebitda|1"
       aGuruFocusItems(62) = "1|ev2rev|EV%252FRevenue|ev2rev|1"
       aGuruFocusItems(63) = "1|forwardPE|Forward%2BP%252FE%2BRatio|N/A|1"
       aGuruFocusItems(64) = "1|RateOfReturn|Forward%2BRate%2Bof%2BReturn|RateOfReturn|1"
       aGuruFocusItems(65) = "1|goodwill2asset|Goodwill%2Bto%2BAsset%2BRatio|goodwill2asset|1"
       aGuruFocusItems(66) = "1|interest_coverage|Interest%2BCoverage|interest_coverage|1"
       aGuruFocusItems(67) = "1|inventory2rev|Inventory%2Bto%2BRevenue|inventory2rev|1"
       aGuruFocusItems(68) = "1|InventoryTurnover|Inventory%2BTurnover|Inventory Turnover|1"
       aGuruFocusItems(69) = "1|ltd2asset|LT%2BDebt%2Bto%2BTotal%2BAsset|ltd2asset|1"
       aGuruFocusItems(70) = "1|pb|P%252FB%2BRatio|pb|1"
       aGuruFocusItems(71) = "1|pe|P%252FE%2BRatio|pe|1"
       aGuruFocusItems(72) = "1|pettm|P%252FE%2BRatio%2528ttm%2529|pettm|1"
       aGuruFocusItems(73) = "1|penri|P%252FE%2Bwithout%2BNRI|penri|1"
       aGuruFocusItems(74) = "1|ps|P%252FS%2BRatio|ps|1"
       aGuruFocusItems(75) = "1|peg|PEG|peg|1"
       aGuruFocusItems(76) = "1|p2tangible_book|Price%2Bto%2BTangible%2BBook|p2tangible_book|1"
       aGuruFocusItems(77) = "1|pfcf|Price-to-Free-Cash-Flow%2Bratio|pfcf|1"
       aGuruFocusItems(78) = "1|quick_ratio|Quick%2BRatio|quick ratio|1"
       aGuruFocusItems(79) = "1|ROA|Return%2Bon%2BAssets|ROA|1"
       aGuruFocusItems(80) = "1|ROC|Return%2Bon%2BCapital|ROC|1"
       aGuruFocusItems(81) = "1|ROC_JOEL|Return%2Bon%2BCapital%2B%2B-%2BJoel%2BGreenblatt|ROC_JOEL|1"
       aGuruFocusItems(82) = "1|ROE|Return%2Bon%2BEquity|ROE|1"
       aGuruFocusItems(83) = "1|ROTA|Return%2Bon%2BTangible%2BAssets|ROTA|1"
       aGuruFocusItems(84) = "1|ROTE|Return%2Bon%2BTangible%2BEquity|ROTE|1"
       aGuruFocusItems(85) = "1|ShillerPE|Shiller%2BP%252FE%2BRatio|N/A|1"
       aGuruFocusItems(86) = "1|Short Ratio|Short%2BRatio|N/A|1"
       aGuruFocusItems(87) = "1|sloanratio|Sloan%2BRatio|sloanratio|1"
       aGuruFocusItems(88) = "1|grossmargin|Gross%2BMargin|Gross Margin|1"
       aGuruFocusItems(89) = "1|netmargin|Net%2BMargin|netmargin|1"
       aGuruFocusItems(90) = "1|operatingmargin|Operating%2BMargin|Operating Margin|1"
       aGuruFocusItems(91) = "1|EPV|Earnings%2BPower%2BValue%2B%2528EPV%2529|EPV|1"
       aGuruFocusItems(92) = "1|GAVA|GAVA|N/A|1"
       aGuruFocusItems(93) = "1|grahamnumber|Graham%2BNumber|grahamnumber|1"
       aGuruFocusItems(94) = "1|iv_dcf_share|Intrinsic%2BValue%2B%2528DCF%2BProjected%2529|iv_dcf_share|1"
       aGuruFocusItems(95) = "1|Intrinsic Value (DCF)|Intrinsic%2BValue%2B%2528DCF%2529|N/A|1"
       aGuruFocusItems(96) = "1|Intrinsic Value (DE)|Intrinsic%2BValue%2B%2528DE%2529|N/A|1"
       aGuruFocusItems(97) = "1|margin_of_safety|Margin%2Bof%2BSafety|N/A|1"
       aGuruFocusItems(98) = "1|medpsvalue|Median%2BP%252FS%2BValue|medpsvalue|1"
       aGuruFocusItems(99) = "1|netcash_per_share|Net%2BCash%2B%2528per%2Bshare%2529|netcash|1"
       aGuruFocusItems(100) = "1|lynchvalue|Peter%2BLynch%2BFair%2BValue|lynchvalue|1"
       aGuruFocusItems(101) = "1|ShortPercentageOfFloat|Short%2BPercentage%2BOf%2BFloat|Short Percentage Of Float|1"
       aGuruFocusItems(102) = "1|dividend_growth_3y|Dividend%2BGrowth%2BRate%2B3y|N/A|1"
       aGuruFocusItems(103) = "1|dividend_growth_5y|Dividend%2BGrowth%2BRate%2B5y|N/A|1"
       aGuruFocusItems(104) = "1|payout|Dividend%2BPayout%2BRatio|payout|1"
       aGuruFocusItems(105) = "1|yield|Dividend%2BYield|N/A|1"
       aGuruFocusItems(106) = "1|yield_on_cost|Yield%2Bon%2BCost|N/A|1"
       aGuruFocusItems(107) = "1|PromotionAndAdvertising|Advertising|PromotionAndAdvertising|1"
       aGuruFocusItems(108) = "1|COGS|Cost%2Bof%2BGoods%2BSold|COGS|1"
       aGuruFocusItems(109) = "1|IS_CreditLossesProvision|Credit%2BLosses%2BProvision|IS_CreditLossesProvision|1"
       aGuruFocusItems(110) = "1|EBITDA|Earnings%2BBefore%2BDepreciation%2Band%2BAmortization|EBITDA|1"
       aGuruFocusItems(111) = "1|eps_basic|EPS%2B%2528Basic%2529|eps_basic|1"
       aGuruFocusItems(112) = "1|eps_diluated|EPS%2B%2528Diluted%2529|eps_diluated|1"
       aGuruFocusItems(113) = "1|IS_FeeRevenueAndOtherIncome|Fees%2Band%2BOther%2BIncome|IS_FeeRevenueAndOtherIncome|1"
       aGuruFocusItems(114) = "1|Gross Profit|Gross%2BProfit|Gross_Profit|1"
       aGuruFocusItems(115) = "1|InterestExpense|Interest%2BExpense|Interest Expense|1"
       aGuruFocusItems(116) = "1|InterestIncome|Interest%2BIncome|InterestIncome|1"
       aGuruFocusItems(117) = "1|Net Income|Net%2BIncome|Net Income|1"
       aGuruFocusItems(118) = "1|Net Income (Continuing Operations)|Net%2BIncome%2B%2528Continuing%2BOperations%2529|Net Income (Continuing Operations)|1"
       aGuruFocusItems(119) = "1|Net Income (Discontinued Operations)|Net%2BIncome%2B%2528Discontinued%2BOperations%2529|Net Income (Discontinued Operations)|1"
       aGuruFocusItems(120) = "1|IS_NetInterestIncome|Net%2BInterest%2BIncome|IS_NetInterestIncome|1"
       aGuruFocusItems(121) = "1|IS_NetInvestmentIncome|Net%2BInvestment%2BIncome|IS_NetInvestmentIncome|1"
       aGuruFocusItems(122) = "1|IS_NonInterestIncome|Non%2BInterest%2BIncome|IS_NonInterestIncome|1"
       aGuruFocusItems(123) = "1|Non Operating Income|Non-Recurring%2BItems|Non Operating Income|1"
       aGuruFocusItems(124) = "1|Operating Income|Operating%2BIncome|Operating Income|1"
       aGuruFocusItems(125) = "1|IS_otherExpense_bank|Other%2BExpenses|IS_otherExpense_bank|1"
       aGuruFocusItems(126) = "1|OtherIncomeExpense|Other%2BIncome%2B%2528Minority%2BInterest%2529|OtherIncomeExpense|1"
       aGuruFocusItems(127) = "1|IS_PolicyAcquisitionExpense|Policy%2BAcquisition%2BExpense|IS_PolicyAcquisitionExpense|1"
       aGuruFocusItems(128) = "1|IS_NetPolicyholderBenefitsAndClaims|Policyholder%2BBenefits%2B%2526%2BClaims|IS_NetPolicyholderBenefitsAndClaims|1"
       aGuruFocusItems(129) = "1|Pretax Income|Pre-Tax%2BIncome|Pretax Income|1"
       aGuruFocusItems(130) = "1|IS_preferred_dividends|Preferred%2Bdividends|IS_preferred_dividends|1"
       aGuruFocusItems(131) = "1|RD|Research%2B%2526%2BDevelopment|Research & Development|1"
       aGuruFocusItems(132) = "1|Revenue|Revenue|Revenue|1"
       aGuruFocusItems(133) = "1|SGA|Selling%252C%2BGeneral%252C%2B%2526%2BAdmin.%2BExpense|SG&A|1"
       aGuruFocusItems(134) = "1|Shares Outstanding|Shares%2BOutstanding|Shares Outstanding|1"
       aGuruFocusItems(135) = "1|SpecialCharges|SpecialCharges|SpecialCharges|1"
       aGuruFocusItems(136) = "1|tax|Tax%2BExpense|tax|1"
       aGuruFocusItems(137) = "1|TaxProvision|Tax%2BProvision|TaxProvision|1"
       aGuruFocusItems(138) = "1|IS_TotalPremiumsEarned|Total%2BPremiums%2BEarned|IS_TotalPremiumsEarned|1"
       aGuruFocusItems(139) = "1|Accts Payable|Accounts%2BPayable|Accts Payable|1"
       aGuruFocusItems(140) = "1|Accts Rec.|Accounts%2BReceivable|Accts Rec.|1"
       aGuruFocusItems(141) = "1|AccumulatedDepreciation|Accumulated%2BDepreciation|AccumulatedDepreciation|1"
       aGuruFocusItems(142) = "1|AdditionalPaidInCapital|Additional%2BPaid-In%2BCapital|AdditionalPaidInCapital|1"
       aGuruFocusItems(143) = "1|BuildingsAndImprovements|Buildings%2BAnd%2BImprovements|BuildingsAndImprovements|1"
       aGuruFocusItems(144) = "1|LongTermCapitalLeaseObligation|Capital%2BLease%2BObligation|LongTermCapitalLeaseObligation|1"
       aGuruFocusItems(145) = "1|Capital Surplus|Capital%2BSurplus|Capital Surplus|1"
       aGuruFocusItems(146) = "1|CashAndCashEquivalents|Cash%2BAnd%2BCash%2BEquivalents|CashAndCashEquivalents|1"
       aGuruFocusItems(147) = "1|BS_CashAndCashEquivalents|Cash%2Band%2Bcash%2Bequivalents|BS_CashAndCashEquivalents|1"
       aGuruFocusItems(148) = "1|CommonStock|Common%2BStock|Common Stock|1"
       aGuruFocusItems(149) = "1|ConstructionInProgress|Construction%2BIn%2BProgress|ConstructionInProgress|1"
       aGuruFocusItems(150) = "1|Short-Term Debt|Current%2BPortion%2Bof%2BLong-Term%2BDebt|Short-Term Debt|1"
       aGuruFocusItems(151) = "1|BS_DeferredPolicyAcquisitionCosts|Deferred%2BPolicy%2BAcquisition%2BCosts|BS_DeferredPolicyAcquisitionCosts|1"
       aGuruFocusItems(152) = "1|BS_EquityInvestments|Equity%2BInvestments|BS_EquityInvestments|1"
       aGuruFocusItems(153) = "1|BS_FixedMaturityInvestment|Fixed%2BMaturity%2BInvestment|BS_FixedMaturityInvestment|1"
       aGuruFocusItems(154) = "1|BS_FuturePolicyBenefits|Future%2BPolicy%2BBenefits|BS_FuturePolicyBenefits|1"
       aGuruFocusItems(155) = "1|GrossPPE|Gross%2BProperty%252C%2BPlant%2Band%2BEquipment|GrossPPE|1"
       aGuruFocusItems(156) = "1|Intangibles|Intangible%2BAssets|Intangibles|1"
       aGuruFocusItems(157) = "1|FinishedGoods|Inventories%252C%2BFinished%2BGoods|FinishedGoods|1"
       aGuruFocusItems(158) = "1|OtherInventories|Inventories%252C%2BOther|OtherInventories|1"
       aGuruFocusItems(159) = "1|RawMaterials|Inventories%252C%2BRaw%2BMaterials%2B%2526%2BComponents|RawMaterials|1"
       aGuruFocusItems(160) = "1|WorkInProcess|Inventories%252C%2BWork%2BIn%2BProcess|WorkInProcess|1"
       aGuruFocusItems(161) = "1|Inventory|Inventory|Inventory|1"
       aGuruFocusItems(162) = "1|LandAndImprovements|Land%2BAnd%2BImprovements|LandAndImprovements|1"
       aGuruFocusItems(163) = "1|Long-Term Debt|Long-Term%2BDebt|Long-Term Debt|1"
       aGuruFocusItems(164) = "1|MarketableSecurities|Marketable%2BSecurities|MarketableSecurities|1"
       aGuruFocusItems(165) = "1|BS_MoneyMarket|Money%2BMarket%2BInvestments|BS_MoneyMarket|1"
       aGuruFocusItems(166) = "1|BS_NetLoan|Net%2BLoan|BS_NetLoan|1"
       aGuruFocusItems(167) = "1|BS_other_assets_Bank|Other%2BAssets|BS_other_assets_Bank|1"
       aGuruFocusItems(168) = "1|Other Current Assets|Other%2BCurrent%2BAssets|Other Current Assets|1"
       aGuruFocusItems(169) = "1|Other Current Liab|Other%2BCurrent%2BLiabilities|Other Current Liab.|1"
       aGuruFocusItems(170) = "1|BS_other_liabilities_bank|Other%2Bliabilities|BS_other_liabilities_bank|1"
       aGuruFocusItems(171) = "1|Other Long-Term Assets|Other%2BLong%2BTerm%2BAssets|Other Long-Term Assets|1"
       aGuruFocusItems(172) = "1|Other Long-Term Liab|Other%2BLong-Term%2BLiabilities|Other Long-Term Liab.|1"
       aGuruFocusItems(173) = "1|Paid-In Capital|Paid-In%2BCapital|Paid-In Capital|1"
       aGuruFocusItems(174) = "1|BS_PolicyholderFunds|Policyholder%2BFunds|BS_PolicyholderFunds|1"
       aGuruFocusItems(175) = "1|Preferred Stock|Preferred%2BStock|Preferred Stock|1"
       aGuruFocusItems(176) = "1|Net PPE|Property%252C%2BPlant%2Band%2BEquipment|Net PPE|1"
       aGuruFocusItems(177) = "1|Retained Earnings|Retained%2BEarnings|Retained Earnings|1"
       aGuruFocusItems(178) = "1|BS_SecuritiesAndInvestments|Securities%2B%2526%2BInvestments|BS_SecuritiesAndInvestments|1"
       aGuruFocusItems(179) = "1|BS_TradingAssets|Short-term%2Binvestments|BS_TradingAssets|1"
       aGuruFocusItems(180) = "1|Total Assets|Total%2BAssets|Total Assets|1"
       aGuruFocusItems(181) = "1|Total Current Assets|Total%2BCurrent%2BAssets|Total Current Assets|1"
       aGuruFocusItems(182) = "1|Total Current Liabilities|Total%2BCurrent%2BLiabilities|Total Current Liabilities|1"
       aGuruFocusItems(183) = "1|BS_TotalDeposits|Total%2BDeposits|BS_TotalDeposits|1"
       aGuruFocusItems(184) = "1|Total Equity|Total%2BEquity|Total Equity|1"
       aGuruFocusItems(185) = "1|Total Liabilities|Total%2BLiabilities|Total Liabilities|1"
       aGuruFocusItems(186) = "1|Treasury Stock|Treasury%2BStock|Treasury Stock|1"
       aGuruFocusItems(187) = "1|BS_UnearnedPremiums|Unearned%2BPremiums|BS_UnearnedPremiums|1"
       aGuruFocusItems(188) = "1|BS_UnpaidLossAndLossReserve|Unpaid%2BLoss%2B%2526%2BLoss%2BReserve|BS_UnpaidLossAndLossReserve|1"
       aGuruFocusItems(189) = "1|Cash Flow_CPEX|Cash%2BFlow%2Bfor%2BCapital%2BExpenditures|Cash Flow_CPEX|1"
       aGuruFocusItems(190) = "1|Dividends|Cash%2BFlow%2Bfor%2BDividends|Dividends|1"
       aGuruFocusItems(191) = "1|Cash Flow from Disc. Op.|Cash%2BFlow%2Bfrom%2BDiscontinued%2BOperations|Cash Flow from Disc. Op.|1"
       aGuruFocusItems(192) = "1|Cash from Financing|Cash%2BFlow%2Bfrom%2BFinancing|Cash from Financing|1"
       aGuruFocusItems(193) = "1|Cash Flow from Investing|Cash%2BFlow%2Bfrom%2BInvesting|Cash Flow from Investing|1"
       aGuruFocusItems(194) = "1|Cash Flow from Operations|Cash%2BFlow%2Bfrom%2BOperations|Cash Flow from Operations|1"
       aGuruFocusItems(195) = "1|Cash Flow from Others|Cash%2BFlow%2Bfrom%2BOthers|Cash Flow from Others|1"
       aGuruFocusItems(196) = "1|CashFromDiscontinuedInvestingActivities|Cash%2BFrom%2BDiscontinued%2BInvesting%2BActivities|CashFromDiscontinuedInvestingActivities|1"
       aGuruFocusItems(197) = "1|CashFromOtherInvestingActivities|Cash%2BFrom%2BOther%2BInvesting%2BActivities|CashFromOtherInvestingActivities|1"
       aGuruFocusItems(198) = "1|ChangeInInventory|Change%2BIn%2BInventory|ChangeInInventory|1"
       aGuruFocusItems(199) = "1|ChangeInReceivables|Change%2BIn%2BReceivables|ChangeInReceivables|1"
       aGuruFocusItems(200) = "1|ChangeInWorkingCapital|Change%2BIn%2BWorking%2BCapital|ChangeInWorkingCapital|1"
       aGuruFocusItems(201) = "1|CumulativeEffectOfAccountingChange|Cumulative%2BEffect%2BOf%2BAccounting%2BChange|CumulativeEffectOfAccountingChange|1"
       aGuruFocusItems(202) = "1|CF_DDA|Depreciation%252C%2BDepletion%2Band%2BAmortization|DDA|1"
       aGuruFocusItems(203) = "1|total_freecashflow|Free%2BCash%2BFlow|total_freecashflow|1"
       aGuruFocusItems(204) = "1|Net Change in Cash|Net%2BChange%2Bin%2BCash|Net Change in Cash|1"
       aGuruFocusItems(205) = "1|NetForeignCurrencyExchangeGainLoss|Net%2BForeign%2BCurrency%2BExchange%2BGain|NetForeignCurrencyExchangeGainLoss|1"
       aGuruFocusItems(206) = "1|NetIncomeFromContinuingOperations|Net%2BIncome%2BFrom%2BContinuing%2BOperations|NetIncomeFromContinuingOperations|1"
       aGuruFocusItems(207) = "1|NetIntangiblesPurchaseAndSale|Net%2BIntangibles%2BPurchase%2BAnd%2BSale|NetIntangiblesPurchaseAndSale|1"
       aGuruFocusItems(208) = "1|Net Issuance of Debt|Net%2BIssuance%2Bof%2BDebt|Net Issuance of Debt|1"
       aGuruFocusItems(209) = "1|Net Issuance of preferred|Net%2BIssuance%2Bof%2BPreferred%2BStock|Net Issuance of preferred|1"
       aGuruFocusItems(210) = "1|Net Issuance of Stock|Net%2BIssuance%2Bof%2BStock|Net Issuance of Stock|1"
       aGuruFocusItems(211) = "1|PurchaseOfPPE|Purchase%2BOf%2BProperty%252C%2BPlant%252C%2BEquipment|PurchaseOfPPE|1"
       aGuruFocusItems(212) = "1|PurchaseOfBusiness|PurchaseOfBusiness|PurchaseOfBusiness|1"
       aGuruFocusItems(213) = "1|PurchaseOfInvestment|PurchaseOfInvestment|PurchaseOfInvestment|1"
       aGuruFocusItems(214) = "1|SaleOfPPE|Sale%2BOf%2BProperty%252C%2BPlant%252C%2BEquipment|SaleOfPPE|1"
       aGuruFocusItems(215) = "1|SaleOfBusiness|SaleOfBusiness|SaleOfBusiness|1"
       aGuruFocusItems(216) = "1|SaleOfInvestment|SaleOfInvestment|SaleOfInvestment|1"
    End Sub

