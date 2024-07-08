Attribute VB_Name = "modGetGuruFocusItem"
'@Lang VBA
Option Explicit
Public aGuruFocusItems2(1 To 250) As Variant
Function smfGetGuruFocusItem(ByVal pTicker As String, _
                             ByVal pItem As Integer, _
                    Optional ByVal pPeriod As String = "TTM", _
                    Optional ByVal pOffset As Integer = 0, _
                    Optional ByVal pError As Variant = "Error", _
                    Optional ByVal pType As Integer = 0) As Variant
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return a data item from GuruFocus
   '-----------------------------------------------------------------------------------------------------------*
   ' 2015.07.08 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2015.07.20 -- Fix "Fiscal Periods" for missing data (i.e. pItem = 0)
   ' 2015.11.19 -- Fix "EBIT per share" extraction because of web page change
   ' 2016.04.09 -- Fix iLQ for web page change (5 quarters instead of 9)
   ' 2016.04.19 -- Fix annual extraction of fiscal periods
   ' 2016.09.02 -- Change extraction due to web page change (">Per Share Data" to "id=""Rf""")
   ' 2017.02.13 -- Change extraction due to web page change ("normal" to "normal_pershare")
   ' 2017.05.03 -- Change "http://" protocol to "https://"
   ' 2017.08.13 -- Miscellaneous label changes
   ' 2017.09.17 -- Miscellaneous label changes
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample of use:
   '
   '    =smfGetGuruFocusItem("MMM",1,"TTM")
   '    =smfGetGuruFocusItem("MMM",1,"A",0)
   '    =smfGetGuruFocusItem("MMM",1,"Q",0)
   '
   '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    vError = pError
       
    '------------------> Array of potential data items
    Dim aLabel As Variant
    If aGuruFocusItems2(1) = "" Then smfLoadGuruFocusItems2
    
    '------------------> Which data period?
    Dim sURL As String, s1 As String
    Dim iFQ As Integer, iLQ As Integer
    sURL = "https://www.gurufocus.com/financials/" & UCase(pTicker)
    If pItem <> 0 Then aLabel = Split(aGuruFocusItems2(pItem), "|")
    Select Case True
       Case UCase(pPeriod) = "L"
            Select Case True
               Case pItem = 0
                    smfGetGuruFocusItem = "Fiscal Period"
               Case Else
                    smfGetGuruFocusItem = aLabel(2)
               End Select
            Exit Function
       Case UCase(pTicker) = "NONE" Or pTicker = ""
            smfGetGuruFocusItem = "--"
            Exit Function
       Case UCase(pPeriod) = "TTM"
            Select Case True
               Case pItem = 0
                    smfGetGuruFocusItem = "TTM"
               Case Else
                    smfGetGuruFocusItem = smfConvertData(smfGetTagContent(sURL, "div", -1, "id=""Rf""", aLabel(0), "yesttm"))
               End Select
            Exit Function
       Case UCase(pPeriod) = "A"
            iFQ = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "td", 1, "id=""Rf""", "/term/per+share+rev/", "ttm"), "normal_pershare", "'"))
            Select Case True
               Case pOffset > (iFQ - 2) Or pOffset < 0
                    smfGetGuruFocusItem = "N/A"
               Case pItem = 0
                    'smfGetGuruFocusItem = smfStrExtr(smfGetTagContent(sURL, "td", iFQ - pOffset - 1, ">Per Share Data", ">Fiscal Period") & "<", "~", "<")
                    smfGetGuruFocusItem = smfStrExtr(smfGetTagContent(sURL, "td", iFQ - pOffset, "id=""Rf""", ">Fiscal Period") & "<", "~", "<")
               Case Else
                    s1 = smfGetTagContent(sURL, "div", -1, "id=""Rf""", aLabel(0), aLabel(1) & (iFQ - pOffset - 1))
                    If Left(s1, 1) = "<" Then
                       smfGetGuruFocusItem = "Premium"
                    Else
                       smfGetGuruFocusItem = smfConvertData(s1)
                       End If
               End Select
            Exit Function
       Case UCase(pPeriod) = "Q"
            iFQ = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "td", 1, "id=""Rf""", "/term/per+share+rev/", "ttm"), "normal_pershare", "'"))
            iLQ = smfConvertData(smfStrExtr(smfGetTagContent(sURL, "td", -1, "id=""Rf""", "/term/per+share+rev/", "ttm", "<tr"), "normal_pershare", "'"))
            Select Case True
               Case pOffset > (iLQ - iFQ) Or pOffset < 0
                    smfGetGuruFocusItem = "N/A"
               Case pItem = 0
                    smfGetGuruFocusItem = smfStrExtr(smfGetTagContent(sURL, "td", iLQ - pOffset + 2, "id=""Rf""", ">Fiscal Period") & "<", "~", "<")
               Case Else
                    s1 = smfGetTagContent(sURL, "div", -1, "id=""Rf""", aLabel(0), aLabel(1) & (iLQ - pOffset))
                    If Left(s1, 1) = "<" Then
                       smfGetGuruFocusItem = "Premium"
                    Else
                       smfGetGuruFocusItem = smfConvertData(s1)
                    End If
               End Select
            Exit Function
       Case Else
            smfGetGuruFocusItem = "Error: Invalid pType parameter -- must be L, TTM, A, Q"
            Exit Function
       End Select

ErrorExit: smfGetGuruFocusItem = vError
                   
   End Function
Sub smfLoadGuruFocusItems2()
   '-----------------------------------------------------------------------------------------------------------*
   ' Array of data items to extract
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.02.15 -- Updated for GuruFocus web page changes
   ' 2017.08.13 -- Various label changes
   '-----------------------------------------------------------------------------------------------------------*
    
    aGuruFocusItems2(1) = "/term/per+share+rev/|normal_pershare|Revenue per Share"
    aGuruFocusItems2(2) = "/term/EBITDA_per_share/|normal_pershare|EBITDA per Share"
    aGuruFocusItems2(3) = "/term/EBIT_per_share/|normal_pershare|EBIT per Share"
    aGuruFocusItems2(4) = "/term/per+share+eps/|normal_pershare|Earnings per Share (diluted)"
    aGuruFocusItems2(5) = "/term/eps_nri/|normal_pershare|eps without NRI"
    aGuruFocusItems2(6) = "/term/per+share_freecashflow/|normal_pershare|Free Cashflow per Share"
    aGuruFocusItems2(7) = "/term/Dividends+Per+Share/|normal_pershare|Dividends Per Share"
    aGuruFocusItems2(8) = "/term/Book+Value+Per+Share/|normal_pershare|Book Value Per Share"
    aGuruFocusItems2(9) = "/term/Tangibles_book_per_share/|normal_pershare|Tangible Book per share"
    aGuruFocusItems2(10) = "/term/Month+End+Stock+Price/|normal_pershare|Month End Stock Price"
    aGuruFocusItems2(11) = "/term/ROE/|normal_ratio|Return on Equity %"
    aGuruFocusItems2(12) = "/term/ROA/|normal_ratio|Return on Assets %"
    aGuruFocusItems2(13) = "/term/ROIC/|normal_ratio|Return on Invested Capital %"
    aGuruFocusItems2(14) = "/term/ROC_JOEL/|normal_ratio|Return on Capital - Joel Greenblatt %"
    aGuruFocusItems2(15) = "/term/deb2equity/|normal_ratio|Debt to Equity"
    aGuruFocusItems2(16) = "/term/grossmargin/|normal_ratio|Gross Margin %"
    aGuruFocusItems2(17) = "/term/operatingmargin/|normal_ratio|Operating Margin %"
    aGuruFocusItems2(18) = "/term/netmargin/|normal_ratio|Net Margin %"
    aGuruFocusItems2(19) = "/term/equity2asset/|normal_ratio|Total Equity to Total Asset"
    aGuruFocusItems2(20) = "/term/debt2asset/|normal_ratio|LT Debt to Total Asset"
    aGuruFocusItems2(21) = "/term/turnover/|normal_ratio|Asset Turnover"
    aGuruFocusItems2(22) = "/term/payout/|normal_ratio|Dividend Payout Ratio"
    aGuruFocusItems2(23) = "/term/DaysSalesOutstanding/|normal_ratio|Days Sales Outstanding"
    aGuruFocusItems2(24) = "/term/DaysPayable/|normal_ratio|Days Accounts Payable"
    aGuruFocusItems2(25) = "/term/DaysInventory/|normal_ratio|Days Inventory"
    aGuruFocusItems2(26) = "/term/CCC/|normal_ratio|Cash Conversion Cycle"
    aGuruFocusItems2(27) = "/term/InventoryTurnover/|normal_ratio|Inventory Turnover"
    aGuruFocusItems2(28) = "/term/cogs2rev/|normal_ratio|COGS to Revenue"
    aGuruFocusItems2(29) = "/term/inventory2rev/|normal_ratio|Inventory to Revenue"
    aGuruFocusItems2(30) = "/term/Revenue/|is|Revenue"
    aGuruFocusItems2(31) = "/term/COGS/|is|Cost of Goods Sold"
    aGuruFocusItems2(32) = "/term/Gross+Profit/|is|Gross Profit"
    aGuruFocusItems2(33) = "/term/grossmargin/|normal|Gross Margin %"
    aGuruFocusItems2(34) = "/term/SGA/|is|Selling, General, & Admin. Expense"
    aGuruFocusItems2(35) = "/term/PromotionAndAdvertising/|is|Advertising"
    aGuruFocusItems2(36) = "/term/RD/|is|Research & Development"
    aGuruFocusItems2(37) = "/term/other_operating_charges/|is|Other Operating Expense"
    aGuruFocusItems2(38) = "/term/Operating+Income/|is|Operating Income"
    aGuruFocusItems2(39) = "/term/operatingmargin/|normal_ratio|Operating Margin %"
    aGuruFocusItems2(40) = "/term/InterestIncome/|is|Interest Income"
    aGuruFocusItems2(41) = "/term/InterestExpense/|is|Interest Expense"
    aGuruFocusItems2(42) = "/term/OtherIncomeExpense/|incent|Other Income (Expense)"
    aGuruFocusItems2(43) = "/term/OtherIncome_minorityinterest/|is|Other Income (Minority Interest)"
    aGuruFocusItems2(44) = "/term/Pretax+Income/|is|Pre-Tax Income"
    aGuruFocusItems2(45) = "/term/TaxProvision/|is|Tax Provision"
    aGuruFocusItems2(46) = "/term/TaxRate/|incent|Tax Rate %"
    aGuruFocusItems2(47) = "/term/Net+Income+%28Continuing+Operations%29/|is|Net Income (Continuing"
    aGuruFocusItems2(48) = "/term/Net+Income+%28Discontinued+Operations%29/|is|Net Income (Discontinued"
    aGuruFocusItems2(49) = "/term/Net+Income/|is|Net Income"
    aGuruFocusItems2(50) = "/term/netmargin/|normal_ratio|Net Margin %"
    aGuruFocusItems2(51) = "/term/IS_preferred_dividends/|is|Preferred dividends"
    aGuruFocusItems2(52) = "/term/eps_basic/|normal|EPS (Basic)"
    aGuruFocusItems2(53) = "/term/eps_diluated/|normal|EPS (Diluted)"
    aGuruFocusItems2(54) = "/term/Shares+Outstanding/|is|Shares Outstanding (Diluted)"
    aGuruFocusItems2(55) = "/term/DDA/|is|Depreciation, Depletion and Amortization"
    aGuruFocusItems2(56) = "/term/EBITDA/|is|EBITDA"
    aGuruFocusItems2(57) = "/term/CashAndCashEquivalents/|bs|Cash And Cash Equivalents"
    aGuruFocusItems2(58) = "/term/MarketableSecurities/|bs|Marketable Securities"
    aGuruFocusItems2(59) = "/term/Cash+and+Equiv/|bs|Cash, Cash Equivalents, Marketable Securities"
    aGuruFocusItems2(60) = "/term/Accts+Rec/|bs|Accounts Receivable"
    aGuruFocusItems2(61) = "/term/RawMaterials/|bs|Inventories, Raw Materials & Components"
    aGuruFocusItems2(62) = "/term/WorkInProcess/|bs|Inventories, Work In Process"
    aGuruFocusItems2(63) = "/term/InventoriesAdjustmentsAllowances/|bs|Inventories, Inventories Adj"
    aGuruFocusItems2(64) = "/term/FinishedGoods/|bs|Inventories, Finished Goods"
    aGuruFocusItems2(65) = "/term/OtherInventories/|bs|Inventories, Other"
    aGuruFocusItems2(66) = "/term/Inventory/|bs|Total Inventories"
    aGuruFocusItems2(67) = "/term/Other+Current+Assets/|bs|Other Current Assets"
    aGuruFocusItems2(68) = "/term/Total+Current+Assets/|bs|Total Current Assets"
    aGuruFocusItems2(69) = "/term/LandAndImprovements/|bs|Land And Improvements"
    aGuruFocusItems2(70) = "/term/BuildingsAndImprovements/|bs|Buildings And Improvements"
    aGuruFocusItems2(71) = "/term/MachineryFurnitureEquipment/|bs|Machinery, Furniture, Equipment"
    aGuruFocusItems2(72) = "/term/ConstructionInProgress/|bs|Construction In Progress"
    aGuruFocusItems2(73) = "/term/GrossPPE/|bs|Gross Property, Plant and Equipment"
    aGuruFocusItems2(74) = "/term/AccumulatedDepreciation/|bs|Accumulated Depreciation"
    aGuruFocusItems2(75) = "/term/Net+PPE/|bs|Property, Plant and Equipment"
    aGuruFocusItems2(76) = "/term/Intangibles/|bs|Intangible Assets"
    aGuruFocusItems2(77) = "/term/Goodwill/|bs|Goodwill"
    aGuruFocusItems2(78) = "/term/Other+Long-Term+Assets/|bs|Other Long Term Assets"
    aGuruFocusItems2(79) = "/term/Total+Assets/|bs|Total Assets"
    aGuruFocusItems2(80) = "/term/AccountsPayable/|bs|Accounts Payable"
    aGuruFocusItems2(81) = "/term/TotalTaxPayable/|bs|Total Tax Payable"
    aGuruFocusItems2(82) = "/term/OtherAccruedExpenses/|bs|Other Accrued Expense"
    aGuruFocusItems2(83) = "/term/Accts+Payable/|bs|Accounts Payable & Accrued Expense"
    aGuruFocusItems2(84) = "/term/Short-Term+Debt/|bs|Current Portion of Long-Term Debt"
    aGuruFocusItems2(85) = "/term/BS_CurrentDeferredLiabilities/|bs|DeferredTaxAndRevenue"
    aGuruFocusItems2(86) = "/term/Other+Current+Liab/|bs|Other Current Liabilities"
    aGuruFocusItems2(87) = "/term/Total+Current+Liabilities/|bs|Total Current Liabilities"
    aGuruFocusItems2(88) = "/term/LongTermDebt/|bs|Long-Term Debt"
    aGuruFocusItems2(89) = "/term/deb2equity/|normal_ratio|Debt to Equity"
    aGuruFocusItems2(90) = "/term/LongTermCapitalLeaseObligation/|bs|Capital Lease Obligation"
    aGuruFocusItems2(91) = "/term/PensionAndRetirementBenefit/|bs|PensionAndRetirementBenefit"
    aGuruFocusItems2(92) = "/term/NonCurrentDeferredLiabilities/|bs|NonCurrent Deferred Liabilities"
    aGuruFocusItems2(93) = "/term/Other+Long-Term+Liab/|bs|Other Long-Term Liabilities"
    aGuruFocusItems2(94) = "/term/Total+Liabilities/|bs|Total Liabilities"
    aGuruFocusItems2(95) = "/term/CommonStock/|bs|Common Stock"
    aGuruFocusItems2(96) = "/term/Preferred+Stock/|bs|Preferred Stock"
    aGuruFocusItems2(97) = "/term/Retained+Earnings/|bs|Retained Earnings"
    aGuruFocusItems2(98) = "/term/accumulated_other_comprehensive_income/|bs|Accumulated other comp"
    aGuruFocusItems2(99) = "/term/AdditionalPaidInCapital/|bs|Additional Paid-In Capital"
    aGuruFocusItems2(100) = "/term/Treasury+Stock/|bs|Treasury Stock"
    aGuruFocusItems2(101) = "/term/Total+Equity/|bs|Total Equity"
    aGuruFocusItems2(102) = "/term/equity2asset/|normal_ratio|Total Equity to Total Asset"
    aGuruFocusItems2(103) = "/term/CF_Net+Income/|cs|Net Income"
    aGuruFocusItems2(104) = "/term/CumulativeEffectOfAccountingChange/|cs|Cumulative Effect Of Acco"
    aGuruFocusItems2(105) = "/term/NetForeignCurrencyExchangeGainLoss/|cs|Net Foreign Currency Exch"
    aGuruFocusItems2(106) = "/term/NetIncomeFromContinuingOperations/|cs|Net Income From Continuing"
    aGuruFocusItems2(107) = "/term/CF_DDA/|cs|Depreciation, Depletion and Amortization"
    aGuruFocusItems2(108) = "/term/ChangeInReceivables/|cs|Change In Receivables"
    aGuruFocusItems2(109) = "/term/ChangeInInventory/|cs|Change In Inventory"
    aGuruFocusItems2(110) = "/term/ChangeInPrepaidAssets/|cs|Change In Prepaid Assets"
    aGuruFocusItems2(111) = "/term/ChangeInPayablesAndAccruedExpense/|cs|Change In Payables And Acc"
    aGuruFocusItems2(112) = "/term/ChangeInWorkingCapital/|cs|Change In Working Capital"
    aGuruFocusItems2(113) = "/term/CF_DeferredTax/|cs|Change In DeferredTax"
    aGuruFocusItems2(114) = "/term/StockBasedCompensation/|cs|Stock Based Compensation"
    aGuruFocusItems2(115) = "/term/Cash+Flow+from+Disc+Op/|cs|Cash Flow from Discontinued Operati"
    aGuruFocusItems2(116) = "/term/Cash+Flow+from+Others/|cs|Cash Flow from Others"
    aGuruFocusItems2(117) = "/term/Cash+Flow+from+Operations/|cs|Cash Flow from Operations"
    aGuruFocusItems2(118) = "/term/PurchaseOfPPE/|cs|Purchase Of Property, Plant, Equipment"
    aGuruFocusItems2(119) = "/term/SaleOfPPE/|cs|Sale Of Property, Plant, Equipment"
    aGuruFocusItems2(120) = "/term/PurchaseOfBusiness/|cs|Purchase Of Business"
    aGuruFocusItems2(121) = "/term/SaleOfBusiness/|cs|Sale Of Business"
    aGuruFocusItems2(122) = "/term/PurchaseOfInvestment/|cs|Purchase Of Investment"
    aGuruFocusItems2(123) = "/term/SaleOfInvestment/|cs|Sale Of Investment"
    aGuruFocusItems2(124) = "/term/NetIntangiblesPurchaseAndSale/|cs|Net Intangibles Purchase And S"
    aGuruFocusItems2(125) = "/term/CashFromDiscontinuedInvestingActivities/|cs|Cash From Discontinu"
    aGuruFocusItems2(126) = "/term/CashFromOtherInvestingActivities/|cs|Cash From Other Investing A"
    aGuruFocusItems2(127) = "/term/Cash+Flow+from+Investing/|cs|Cash Flow from Investing"
    aGuruFocusItems2(128) = "/term/Issuance_of_Stock/|cs|Issuance of Stock"
    aGuruFocusItems2(129) = "/term/Repurchase_of_Stock/|cs|Repurchase of Stock"
    aGuruFocusItems2(130) = "/term/Net+Issuance+of+preferred/|cs|Net Issuance of Preferred Stock"
    aGuruFocusItems2(131) = "/term/Net+Issuance+of+Debt/|cs|Net Issuance of Debt"
    aGuruFocusItems2(132) = "/term/Dividends/|cs|Cash Flow for Dividends"
    aGuruFocusItems2(133) = "/term/Other+Financing/|cs|Other Financing"
    aGuruFocusItems2(134) = "/term/Cash+from+Financing/|cs|Cash Flow from Financing"
    aGuruFocusItems2(135) = "/term/Net+Change+in+Cash/|cs|Net Change in Cash"
    aGuruFocusItems2(136) = "/term/Cash+Flow_CPEX/|cs|Capital Expenditure"
    aGuruFocusItems2(137) = "/term/total_freecashflow/|cs|Free Cash Flow"
    aGuruFocusItems2(138) = "/term/pettm/|normal_vratio|PE Ratio(ttm)"
    aGuruFocusItems2(139) = "/term/pb/|normal_vratio|Price to Book"
    aGuruFocusItems2(140) = "/term/p2tangible_book/|normal_vratio|Price to Tangible Book"
    aGuruFocusItems2(141) = "/term/pfcf/|normal_vratio|Price-to-Free-Cash-Flow ratio"
    aGuruFocusItems2(142) = "/term/ps/|normal_vratio|PS Ratio"
    aGuruFocusItems2(143) = "/term/peg/|normal_vratio|PEG Ratio"
    aGuruFocusItems2(144) = "/term/ev2rev/|normal_vratio|EV-to-Revenue"
    aGuruFocusItems2(145) = "/term/ev2ebitda/|normal_vratio|EV-to-EBITDA"
    aGuruFocusItems2(146) = "/term/ev2ebit/|normal_vratio|EV-to-EBIT"
    aGuruFocusItems2(147) = "/term/earning_yield_greenblatt/|normal_vratio|Earnings Yield (Joel Greenblatt"
    aGuruFocusItems2(148) = "/term/RateOfReturn/|normal_vratio|Forward Rate of Return"
    aGuruFocusItems2(149) = "/term/ShillerPE/|normal_vratio|Shiller PE Ratio"
    aGuruFocusItems2(150) = "/term/mktcap/|normal_vq|Market Cap"
    aGuruFocusItems2(151) = "/term/ev/|normal_vq|Enterprise Value"
    aGuruFocusItems2(152) = "/term/Month+End+Stock+Price/|normal_pershare|Month End Stock Price"
    aGuruFocusItems2(153) = "/term/netcash/|normal_vq|Net Cash (per share)"
    aGuruFocusItems2(154) = "/term/NCAV_real/|normal_vq|Net Current Asset Value (per share)"
    aGuruFocusItems2(155) = "/term/iv_dcf_share/|normal_vq|Projected FCF (per share)"
    aGuruFocusItems2(156) = "/term/medpsvalue/|normal_vq|Median PS (per share)"
    aGuruFocusItems2(157) = "/term/lynchvalue/|normal_vq|Peter Lynch Fair Value (per share)"
    aGuruFocusItems2(158) = "/term/grahamnumber/|normal_vq|Graham Number (per share)"
    aGuruFocusItems2(159) = "/term/EPV/|normal_vq|Earning Power Value (per share)"
    aGuruFocusItems2(160) = "/term/zscore/|normal_vq|Altman Z-Score"
    aGuruFocusItems2(161) = "/term/fscore/|normal_vq|Piotroski F-Score"
    aGuruFocusItems2(162) = "/term/mscore/|normal_vq|Beneish M-Score"
    aGuruFocusItems2(163) = "/term/sloanratio/|normal_vq|Sloan Ratio (%)"
    aGuruFocusItems2(164) = "/term/price_high/|normal_vq|Highest Stock Price"
    aGuruFocusItems2(165) = "/term/price_low/|normal_vq|Lowest Stock Price"
    aGuruFocusItems2(166) = "/term/total_buyback_3y/|normal|Shares Buyback Ratio (%)"
    aGuruFocusItems2(167) = "/term/growth_per_share_rev/|normal|YoY Rev. per Sh. Growth (%)"
    aGuruFocusItems2(168) = "/term/growth_per_share_eps/|normal|YoY EPS Growth (%)"
    aGuruFocusItems2(169) = "/term/growth_per_share_ebitda/|normal|YoY EBITDA Growth (%)"
    aGuruFocusItems2(170) = "/term/editda_5y_growth/|normal_vq|EBITDA 5-Y Growth (%)"
    aGuruFocusItems2(171) = "/term/shares_basic/|normal_vq|Shares Outstanding (Basic)"
    aGuruFocusItems2(172) = "/term/BS_share/|normal_vq|Shares Outstanding"
    aGuruFocusItems2(173) = "/term/IS_NetInterestIncome/|is|Net Interest Income"
    aGuruFocusItems2(174) = "/term/IS_NonInterestIncome/|is|Non Interest Income"
    aGuruFocusItems2(175) = "/term/IS_CreditLossesProvision/|is|Credit Losses Provision"
    aGuruFocusItems2(176) = "/term/SpecialCharges/|is|Special Charges"
    aGuruFocusItems2(177) = "/term/IS_otherExpense_bank/|is|Other Noninterest Expense"
    aGuruFocusItems2(178) = "/term/BS_CashAndCashEquivalents/|bs|Cash and cash equivalents"
    aGuruFocusItems2(179) = "/term/BS_MoneyMarket/|bs|Money Market Investments"
    aGuruFocusItems2(180) = "/term/BS_NetLoan/|bs|Net Loan"
    aGuruFocusItems2(181) = "/term/BS_SecuritiesAndInvestments/|bs|Securities & Investments"
    aGuruFocusItems2(182) = "/term/BS_DeferredPolicyAcquisitionCosts/|bs|Deferred Policy Acquisitio"
    aGuruFocusItems2(183) = "/term/BS_other_assets_Bank/|bs|Other Assets"
    aGuruFocusItems2(184) = "/term/BS_TotalDeposits/|bs|Total Deposits"
    aGuruFocusItems2(185) = "/term/BS_other_liabilities_bank/|bs|Other liabilities"

    End Sub
