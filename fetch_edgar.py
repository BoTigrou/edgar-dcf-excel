import sys
import argparse
import pandas as pd
import re
import datetime
from edgar import Company, set_identity


def clean_concept(c):
    if not isinstance(c, str):
        return ""
    return re.sub(r'^[^_]+_', '', c).strip()


def is_valid_statement_row(row):
    """Return True only for primary (non-abstract, non-breakdown, non-dimensioned) statement rows."""
    if row.get('abstract') == True or row.get('is_abstract') == True:
        return False
    if row.get('is_breakdown') == True:
        return False
    if row.get('dimension') == True:
        return False
    dim_axis = row.get('dimension_axis')
    if pd.notna(dim_axis) and str(dim_axis).strip() != '':
        return False
    dim_member = row.get('dimension_member')
    if pd.notna(dim_member) and str(dim_member).strip() != '':
        return False
    return True


def map_raw_concept(concept_name):
    clean = clean_concept(concept_name).lower()

    if clean in ['operationsandsupportexpense', 'operationsandsupport', 'operationssupportexpense', 'operationssupport']:
        return 'OperationSupport'

    if clean in ['sellingandmarketingexpense', 'sellingandmarketing', 'salesandmarketingexpense', 'salesandmarketing']:
        return 'SalesMarketing'

    if clean in ['researchanddevelopmentexpense', 'researchanddevelopment']:
        return 'ResearchDevelopment'

    if clean in ['generalandadministrativeexpense', 'generalandadministrative', 'sellinggeneralandadministrativeexpense', 'sellinggeneralandadministrative']:
        return 'GeneralAdministrativeRaw'

    if clean in [
        'depreciationdepletionandamortization',
        'depreciationandamortization',
        'depreciationdepletionandamortizationexpense',
        'amortizationofintangibleassets',
        'restructuringandothercharges',
    ]:
        if clean in ['depreciationdepletionandamortization', 'depreciationandamortization', 'depreciationdepletionandamortizationexpense']:
            return 'DepreciationAmortizationIS'

    if clean in ['interestexpense', 'interestexpensenonoperating', 'interestexpenseonlongtermdebt', 'interestexpenseondebt']:
        return 'InterestExpenseRaw'

    if clean in ['interestincome', 'investmentincomeinterest', 'interestincomeoperating', 'investmentincomeinterestanddividend', 'interestanddividendincomeoperating']:
        return 'InterestIncomeRaw'

    if clean in ['othernonoperatingincomeexpense', 'nonoperatingincomeexpense', 'otheroperatingincomeexpense', 'otherincomeexpense']:
        return 'OtherIncome'

    if clean in ['incometaxexpensebenefit', 'incometaxexpensebenefitcurrent', 'provisionforincometaxes']:
        return 'TaxProvision'

    if clean in ['incomelossfromequitymethodinvestments', 'equitymethodinvestmentrealizedgainloss']:
        return 'EquityInvestmentIncome'

    if clean in ['netincomelossattributabletononcontrollinginterest', 'netincomelossattributabletononcontrollinginterests']:
        return 'NCI_Income'

    if clean in ['netincomeloss', 'netincomelossavailabletocommonstockholdersbasic', 'profitloss']:
        return 'NetIncome'

    if clean in ['operatingincomeloss', 'operatingincome']:
        return 'OperatingIncomeRaw'

    if 'revenue' in clean or 'salesnet' in clean or clean == 'revenues':
        if not any(x in clean for x in ['cost', 'expense', 'deferred', 'accrued', 'receivable', 'liability', 'remainingperformanceobligation']):
            return 'Revenue'

    if any(x in clean for x in ['costofgoods', 'costofrevenue', 'costofgoodsandservice', 'costofsales']):
        if not any(x in clean for x in ['operating', 'operatingexpenses']):
            return 'COGS'

    return None


def map_raw_bs_concept(concept_name):
    clean = clean_concept(concept_name)
    clean_lower = clean.lower()

    if clean_lower in ['cashandcashequivalentsatcarryingvalue', 'cash', 'cashequivalentsatcarryingvalue']:
        return 'CashEquivalents'
    if clean_lower in ['marketablesecuritiescurrent', 'shortterminvestments', 'availableforsalesecuritiesdebtsecuritiescurrent', 'otherinvestmentsandsecuritiesatcost']:
        return 'ShortTermInvestments'
    if clean_lower in ['restrictedcashandcashequivalentsatcarryingvalue', 'restrictedcashandinvestmentscurrent', 'restrictedcashcurrent', 'restrictedcash']:
        return 'RestrictedCashCurrent'
    if clean_lower in [
        'accountsreceivablesnetcurrent',
        'accountsreceivableafterallowanceforcreditlosscurrent',
        'accountsreceivablesnet',
        'accountsreceivablenetcurrent',
        'accountsreceivablecurrent',
        'accountsreceivable',
        'receivablesnetcurrent',
    ]:
        return 'AccountsReceivable'
    if clean_lower in [
        'prepaidexpenseandotherassetscurrent',
        'prepaidexpensecurrent',
        'prepaidexpenseandotherassets',
        'inventorynet',
        'inventories',
        'otherassetscurrent',
        'contractwithcustomerassetnetcurrent',
        'incometaxesreceivable',
        'nontradereceivablescurrent',
    ]:
        return 'PrepaidExpenses'
    if clean_lower in ['restrictedcashandcashequivalentsnoncurrent', 'restrictedcashnoncurrent', 'restrictedcashandinvestmentsnoncurrent']:
        return 'RestrictedCashNoncurrent'
    if clean_lower in ['restrictedinvestmentsnoncurrent', 'restrictedinvestments']:
        return 'RestrictedInvestmentsNoncurrent'
    if clean_lower in [
        'investmentsnoncurrent',
        'marketablesecuritiesnoncurrent',
        'marketableandnonmarketableinvestments',
        'otherlongterminvestments',
        'availableforsalesecuritiesdebtsecuritiesnoncurrent',
        'longterminvestments',
    ]:
        return 'InvestmentsNoncurrent'
    if clean_lower in ['equitymethodinvestments']:
        return 'EquityMethodInvestments'
    if clean_lower in [
        'propertyplantandequipmentnet',
        'propertyplantandequipmentgross',
        'propertyplantandequipmentandfinanceleaserightofuseassetafteraccumulateddepreciationandamortization',
    ]:
        return 'PP_E'
    if clean_lower in ['operatingleaserightofuseasset', 'financeleaserightofuseasset']:
        return 'LeaseAssets'
    if clean_lower in ['intangibleassetsnetexcludinggoodwill', 'intangibleassetsnet', 'finitelivedintangibleassetsnet', 'indefinitelivedintangibleassetsexcludinggoodwill']:
        return 'Intangibles'
    if clean_lower in ['goodwill']:
        return 'Goodwill'
    if clean_lower in ['otherassetsnoncurrent', 'otherassets', 'deferredincometaxassetsnet', 'deferredtaxassetsnet', 'contractwithcustomerassetnetnoncurrent']:
        return 'OtherAssets'
    if clean_lower in ['accountspayablecurrent', 'accountspayable', 'accountspayabletradecurrent']:
        return 'AccountsPayable'
    if clean_lower in [
        'shorttermborrowings',
        'longtermdebtandcapitalleaseobligationscurrent',
        'longtermdebtcurrent',
        'shorttermdebt',
        'currentportionoflongtermdebt',
        'commercialpaper',
        'debtcurrent',
        'notespayablecurrent',
    ]:
        return 'ShortTermDebt'
    if clean_lower in ['operatingleaseliabilitycurrent', 'financeleaseliabilitycurrent']:
        return 'LeaseLiabilitiesCurrent'
    if clean_lower in [
        'accruedliabilitiescurrent',
        'accruedexpensesandotherliabilitiescurrent',
        'accruedliabilitiesandotherliabilitiescurrent',
        'insurancereservecurrent',
        'employeerelatedliabilitiescurrent',
        'otherliabilitiescurrent',
        'contractwithcustomerliabilitycurrent',
        'deferredrevenuecurrent',
        'accruedrevenueshare',
        'accruedincometaxescurrent',
        'taxespayablecurrent',
    ]:
        return 'AccruedLiabilitiesCurrent'
    if clean_lower in ['longterminsurancereserves', 'insurancereservesnoncurrent', 'insurancereservenoncurrent']:
        return 'LongTermInsuranceReserves'
    if clean_lower in ['longtermdebtexcludingcurrentmaturities', 'longtermdebt', 'longtermdebtnoncurrent', 'longtermdebtandcapitalleaseobligations', 'debtnoncurrent']:
        return 'LongTermDebt'
    if clean_lower in ['operatingleaseliabilitynoncurrent', 'financeleaseliabilitynoncurrent']:
        return 'LeaseLiabilitiesNoncurrent'
    if clean_lower in [
        'otherliabilitiesnoncurrent',
        'otherliabilities',
        'deferredincometaxliabilitiesnet',
        'contractwithcustomerliabilitynoncurrent',
        'deferredrevenuenoncurrent',
        'accruedincometaxesnoncurrent',
        'liabilityforuncertaintaxpositionsnoncurrent',
    ]:
        return 'OtherLiabilitiesNoncurrent'
    if clean_lower in [
        'redeemablenoncontrollinginterestequitycarryingamount',
        'temporaryequitypreferredstockdividendobligation',
        'preferredstockdividendobligation',
        'temporaryequitycarryingamountincludingportionattributabletononcontrollinginterests',
        'redeemablenoncontrollinginterestequitycommoncarryingamount',
        'minorityinterest',
        'noncontrollinginterest',
        'nonredeemablenoncontrollinginterest',
    ]:
        return 'RedeemableNCI'
    if clean_lower in [
        'additionalpaidincapital',
        'additionalpaidincapitalcommonstock',
        'commonstockvalue',
        'commonstocksincludingadditionalpaidincapital',
        'preferredstockvalue',
        'convertiblepreferredstocknonredeemableorredeemableissueroptionvalue',
        'commonstockvalueoutstanding',
    ]:
        return 'APIC'
    if clean_lower in ['accumulatedothercomprehensiveincomelossnetoftax']:
        return 'AccumOCI'
    if clean_lower in ['retainedearningsaccumulateddeficit']:
        return 'AccumDeficit'

    return None


def find_year_column(df, year):
    year_str = str(year)
    for col in df.columns:
        if year_str in str(col) and not any(x in str(col).lower() for x in ['parenthetical', 'depth', 'is_abstract', 'is_total', 'section', 'confidence']):
            return col
    return None


def get_concept_values(df, concept_names, year_cols):
    res = {}
    for col in year_cols:
        val = 0.0
        found = False
        if col in df.columns:
            for name in concept_names:
                names_to_try = [name, name.replace('us-gaap_', ''), 'us-gaap_' + name.replace('us-gaap_', '')]
                for n in names_to_try:
                    if n in df.index:
                        row_val = df.loc[n, col]
                        if pd.notna(row_val):
                            if isinstance(row_val, pd.Series):
                                row_val = row_val.iloc[0]
                            val = float(row_val)
                            found = True
                            break
                if found:
                    break
        res[col] = val
    return res


def extract_shares_from_xbrl_facts(xbrl, target_years):
    """Extract diluted (or basic) weighted average shares for target_years from XBRL facts."""
    shares_by_year = {}
    try:
        df_facts = xbrl.facts.to_dataframe()
        dur = df_facts[df_facts['period_type'] == 'duration'].copy()
        if dur.empty:
            return shares_by_year

        concepts_priority = [
            'us-gaap:WeightedAverageNumberOfDilutedSharesOutstanding',
            'us-gaap:WeightedAverageNumberOfSharesOutstandingBasic',
        ]
        for y in target_years:
            y_str = str(y)
            # Match period_end in year y (or early Jan/Feb of y+1 for 52/53-week fiscal years)
            dur_y = dur[dur['period_end'].astype(str).str.startswith(y_str)]
            if dur_y.empty:
                continue
            for c_name in concepts_priority:
                c_rows = dur_y[dur_y['concept'] == c_name]
                if c_rows.empty:
                    continue
                non_dim = c_rows[c_rows['is_dimensioned'] == False]
                if not non_dim.empty:
                    val = float(non_dim['numeric_value'].max())
                    if val > 0:
                        shares_by_year[y] = val
                        break
                else:
                    # Multi-class shares without undimensioned total: sum latest period_end per class
                    dim_axis = c_rows[c_rows['dimension_axis'].astype(str).str.contains('ClassOfStock', na=False)]
                    if not dim_axis.empty:
                        val = float(dim_axis.groupby('dimension_member_label')['numeric_value'].max().sum())
                        if val > 0:
                            shares_by_year[y] = val
                            break
    except Exception:
        pass
    return shares_by_year


def get_latest_quarterly_data(company):
    latest_cash = 0.0
    latest_debt = 0.0
    latest_shares = 1.0  # default fallback

    try:
        latest_filing = company.get_filings(form=["10-K", "10-Q"])[0]
        xbrl = latest_filing.xbrl()
        df_facts = xbrl.facts.to_dataframe()
        df_non_dim = df_facts[df_facts['is_dimensioned'] == False]

        # 1. Latest BS Date from cash
        cash_rows = df_non_dim[
            (df_non_dim['concept'] == 'us-gaap:CashAndCashEquivalentsAtCarryingValue') &
            (df_non_dim['period_instant'].notna())
        ]
        if cash_rows.empty:
            latest_bs_date = df_non_dim[df_non_dim['period_type'] == 'instant']['period_instant'].max()
        else:
            latest_bs_date = cash_rows['period_instant'].max()

        df_bs_latest = df_non_dim[
            (df_non_dim['period_type'] == 'instant') &
            (df_non_dim['period_instant'] == latest_bs_date)
        ]

        # Cash & ST Investments
        cash_concepts = [
            'us-gaap:CashAndCashEquivalentsAtCarryingValue',
            'us-gaap:CashEquivalentsAtCarryingValue',
            'us-gaap:Cash',
            'us-gaap:ShortTermInvestments',
            'us-gaap:MarketableSecuritiesCurrent',
        ]
        for concept in cash_concepts:
            rows = df_bs_latest[df_bs_latest['concept'] == concept]
            if not rows.empty:
                latest_cash += float(rows.iloc[0]['numeric_value'])

        # Debt
        st_debt_concepts = [
            'us-gaap:ShortTermBorrowings',
            'us-gaap:CommercialPaper',
            'us-gaap:LongTermDebtAndCapitalLeaseObligationsCurrent',
            'us-gaap:LongTermDebtCurrent',
        ]
        lt_debt_concepts = [
            'us-gaap:LongTermDebtNoncurrent',
            'us-gaap:LongTermDebtExcludingCurrentMaturities',
            'us-gaap:LongTermDebt',
            'us-gaap:LongTermDebtAndCapitalLeaseObligations',
        ]

        current_debt = 0.0
        for concept in st_debt_concepts:
            rows = df_bs_latest[df_bs_latest['concept'] == concept]
            if not rows.empty:
                current_debt = max(current_debt, float(rows.iloc[0]['numeric_value']))

        noncurrent_debt = 0.0
        for concept in lt_debt_concepts:
            rows = df_bs_latest[df_bs_latest['concept'] == concept]
            if not rows.empty:
                noncurrent_debt = max(noncurrent_debt, float(rows.iloc[0]['numeric_value']))

        latest_debt = current_debt + noncurrent_debt

        # Shares Outstanding
        shares_concepts = [
            'dei:EntityCommonStockSharesOutstanding',
            'us-gaap:CommonStockSharesOutstanding',
            'us-gaap:CommonStockSharesIssued',
        ]
        for concept in shares_concepts:
            rows = df_non_dim[(df_non_dim['concept'] == concept) & (df_non_dim['period_instant'].notna())]
            if not rows.empty:
                latest_row = rows.sort_values(by='period_instant').iloc[-1]
                val = float(latest_row['numeric_value'])
                if val > 0:
                    latest_shares = val
                    break
        if not latest_shares or latest_shares == 1.0:
            # Check dimensioned dei:EntityCommonStockSharesOutstanding (e.g. multi-class GOOGL)
            dim_dei = df_facts[(df_facts['concept'] == 'dei:EntityCommonStockSharesOutstanding') & (df_facts['period_instant'].notna())]
            if not dim_dei.empty:
                max_date = dim_dei['period_instant'].max()
                latest_shares = float(dim_dei[dim_dei['period_instant'] == max_date]['numeric_value'].sum())
        if not latest_shares or latest_shares == 1.0:
            durations = df_non_dim[df_non_dim['period_type'] == 'duration']
            dil_shares_rows = durations[durations['concept'] == 'us-gaap:WeightedAverageNumberOfDilutedSharesOutstanding']
            if not dil_shares_rows.empty:
                latest_row = dil_shares_rows.sort_values(by='period_end').iloc[-1]
                latest_shares = float(latest_row['numeric_value'])

    except Exception as e:
        print(f"Warning: Failed to fetch latest quarterly data: {str(e)}")

    return latest_cash, latest_debt, latest_shares


def main():
    parser = argparse.ArgumentParser(description="Fetch SEC EDGAR historical financials for DCF model.")
    parser.add_argument("pos_ticker", nargs="?", type=str, help="Stock ticker positional (optional)")
    parser.add_argument("--ticker", type=str, default=None, help="Stock ticker (e.g. UBER)")
    parser.add_argument("--email", type=str, default="dcf_user@example.com", help="SEC identity email")
    args = parser.parse_args()

    ticker = (args.ticker or args.pos_ticker or "AVGO").upper()
    set_identity(args.email)

    print(f"Fetching EDGAR data for {ticker}...")
    try:
        company = Company(ticker)
        print(f"Company: {company.name} (CIK: {company.cik})")
    except Exception as e:
        print(f"Error fetching company details: {str(e)}")
        sys.exit(1)

    filings = company.get_filings(form="10-K")
    if len(filings) == 0:
        print("Error: No 10-K filings found for this company.")
        sys.exit(1)

    latest_year = datetime.datetime.now().year - 1

    try:
        latest_filing = filings[0]
        latest_xbrl = latest_filing.xbrl()

        latest_is_stmt = None
        for stmt_dict in latest_xbrl.statements.statements:
            if stmt_dict.get('type') == 'IncomeStatement':
                latest_is_stmt = latest_xbrl.statements.get(stmt_dict['role_name'])
                break

        if not latest_is_stmt:
            print("Error: Could not find raw income statement in the latest 10-K.")
            sys.exit(1)

        df_latest_is = latest_is_stmt.to_dataframe(view='standard')

        latest_years = []
        for col in df_latest_is.columns:
            match = re.search(r'\d{4}', str(col))
            if match:
                latest_years.append(int(match.group()))
        latest_years = sorted(list(set(latest_years)))
        if len(latest_years) > 0:
            latest_year = latest_years[-1]

    except Exception as e:
        print(f"Error reading latest 10-K: {str(e)}")
        sys.exit(1)

    target_years = [latest_year - 4, latest_year - 3, latest_year - 2, latest_year - 1, latest_year]
    print(f"Target fiscal years to extract: {target_years}")

    is_data = {
        'Revenue': {y: 0.0 for y in target_years},
        'COGS': {y: 0.0 for y in target_years},
        'OperationSupport': {y: 0.0 for y in target_years},
        'SalesMarketing': {y: 0.0 for y in target_years},
        'ResearchDevelopment': {y: 0.0 for y in target_years},
        'GeneralAdministrativeRaw': {y: 0.0 for y in target_years},
        'DepreciationAmortizationIS': {y: 0.0 for y in target_years},
        'OperatingIncomeRaw': {y: 0.0 for y in target_years},
        'InterestExpenseRaw': {y: 0.0 for y in target_years},
        'InterestIncomeRaw': {y: 0.0 for y in target_years},
        'OtherIncome': {y: 0.0 for y in target_years},
        'TaxProvision': {y: 0.0 for y in target_years},
        'EquityInvestmentIncome': {y: 0.0 for y in target_years},
        'NCI_Income': {y: 0.0 for y in target_years},
        'NetIncome': {y: 0.0 for y in target_years},
        'DilutedShares': {y: 1.0 for y in target_years},
    }
    eps_diluted_by_year = {y: 0.0 for y in target_years}

    filings_to_load = [filings[0]]
    if len(filings) > 2:
        filings_to_load.append(filings[2])
    elif len(filings) > 1:
        filings_to_load.append(filings[1])

    for filing in filings_to_load:
        try:
            print(f"Parsing filing from {filing.filing_date} (Form {filing.form})...")
            xbrl = filing.xbrl()
            is_stmt = None
            for stmt_dict in xbrl.statements.statements:
                if stmt_dict.get('type') == 'IncomeStatement':
                    is_stmt = xbrl.statements.get(stmt_dict['role_name'])
                    break
            if not is_stmt:
                continue

            df_is = is_stmt.to_dataframe(view='standard').reset_index()
            facts_shares = extract_shares_from_xbrl_facts(xbrl, target_years)

            for y in target_years:
                if is_data['DilutedShares'][y] == 1.0 and y in facts_shares:
                    is_data['DilutedShares'][y] = facts_shares[y]

                col_name = find_year_column(df_is, y)
                if col_name:
                    for idx, row in df_is.iterrows():
                        if not is_valid_statement_row(row):
                            continue
                        concept = row.get('concept')
                        if pd.notna(concept) and isinstance(concept, str):
                            clean_c = clean_concept(concept).lower()
                            if clean_c == 'earningspersharediluted' and eps_diluted_by_year[y] == 0.0:
                                v_eps = row.get(col_name)
                                if pd.notna(v_eps):
                                    if isinstance(v_eps, pd.Series):
                                        v_eps = v_eps.iloc[0]
                                    eps_diluted_by_year[y] = float(v_eps)

                            mapped_key = map_raw_concept(concept)
                            if mapped_key:
                                val = row.get(col_name)
                                if pd.notna(val):
                                    if isinstance(val, pd.Series):
                                        val = val.iloc[0]
                                    if is_data[mapped_key][y] == 0.0:
                                        is_data[mapped_key][y] = float(val)

                    if is_data['DilutedShares'][y] == 1.0:
                        shares_concepts = ['WeightedAverageNumberOfDilutedSharesOutstanding', 'WeightedAverageNumberOfSharesOutstandingBasic']
                        for s_c in shares_concepts:
                            shares_row = df_is[df_is['concept'].str.contains(s_c, na=False)]
                            if not shares_row.empty:
                                for _, s_r in shares_row.iterrows():
                                    if not is_valid_statement_row(s_r):
                                        continue
                                    shares_val = s_r.get(col_name)
                                    if pd.notna(shares_val):
                                        if isinstance(shares_val, pd.Series):
                                            shares_val = shares_val.iloc[0]
                                        if float(shares_val) > 0:
                                            is_data['DilutedShares'][y] = float(shares_val)
                                            break
                            if is_data['DilutedShares'][y] != 1.0:
                                break

                    # Fallback: compute diluted shares from NetIncome / DilutedEPS if needed
                    if is_data['DilutedShares'][y] == 1.0 and eps_diluted_by_year[y] != 0.0 and is_data['NetIncome'][y] != 0.0:
                        is_data['DilutedShares'][y] = abs(is_data['NetIncome'][y] / eps_diluted_by_year[y])

        except Exception as e:
            print(f"Warning: Failed to parse raw IS from filing {filing.accession_no}: {str(e)}")

    # Track which years already had explicit D&A on the face of the Income Statement
    has_raw_is_da = {y: (is_data['DepreciationAmortizationIS'][y] != 0.0) for y in target_years}

    # Retrieve standardized Balance Sheet and Cash Flow, plus raw Balance Sheet & Cash Flow
    try:
        print("Retrieving standardized Balance Sheet and Cash Flow...")
        bs_stmt = company.balance_sheet(periods=5)
        cf_stmt = company.cashflow_statement(periods=5)

        df_bal = bs_stmt.to_dataframe()
        df_cf = cf_stmt.to_dataframe()

        selected_cols_std = [f"FY {y}" for y in target_years]

        bs_keys = {
            'CashEquivalents': ['CashAndCashEquivalentsAtCarryingValue', 'CashCashEquivalentsRestrictedCashAndCashEquivalents'],
            'ShortTermInvestments': ['ShortTermInvestments', 'MarketableSecuritiesCurrent'],
            'RestrictedCashCurrent': ['RestrictedCashAndCashEquivalentsAtCarryingValue', 'RestrictedCashAndInvestmentsCurrent'],
            'AccountsReceivable': ['AccountsReceivableNetCurrent', 'AccountsReceivableAfterAllowanceForCreditLossCurrent', 'AccountsReceivableNet'],
            'PrepaidExpenses': ['PrepaidExpenseAndOtherAssetsCurrent', 'PrepaidExpenseCurrent', 'PrepaidExpenseAndOtherAssets'],
            'RestrictedCashNoncurrent': ['RestrictedCashAndCashEquivalentsNoncurrent'],
            'RestrictedInvestmentsNoncurrent': ['RestrictedInvestmentsNoncurrent'],
            'InvestmentsNoncurrent': ['InvestmentsNoncurrent', 'MarketableSecuritiesNoncurrent'],
            'EquityMethodInvestments': ['EquityMethodInvestments'],
            'PP_E': ['PropertyPlantAndEquipmentNet', 'PropertyPlantAndEquipmentGross'],
            'LeaseAssets': ['OperatingLeaseRightOfUseAsset'],
            'Intangibles': ['IntangibleAssetsNetExcludingGoodwill', 'IntangibleAssetsNet'],
            'Goodwill': ['Goodwill'],
            'OtherAssets': ['OtherAssetsNoncurrent', 'OtherAssets'],
            'AccountsPayable': ['AccountsPayableCurrent', 'AccountsPayable'],
            'ShortTermDebt': ['ShortTermBorrowings', 'LongTermDebtAndCapitalLeaseObligationsCurrent', 'LongTermDebtCurrent', 'ShortTermDebt', 'CurrentPortionOfLongTermDebt'],
            'LeaseLiabilitiesCurrent': ['OperatingLeaseLiabilityCurrent'],
            'AccruedLiabilitiesCurrent': ['AccruedLiabilitiesCurrent', 'AccruedExpensesAndOtherLiabilitiesCurrent'],
            'LongTermInsuranceReserves': ['LongTermInsuranceReserves', 'InsuranceReservesNoncurrent'],
            'LongTermDebt': ['LongTermDebtExcludingCurrentMaturities', 'LongTermDebt', 'LongTermDebtNoncurrent'],
            'LeaseLiabilitiesNoncurrent': ['OperatingLeaseLiabilityNoncurrent'],
            'OtherLiabilitiesNoncurrent': ['OtherLiabilitiesNoncurrent', 'OtherLiabilities'],
            'RedeemableNCI': ['RedeemableNoncontrollingInterestEquityCarryingAmount'],
            'APIC': ['AdditionalPaidInCapital'],
            'AccumOCI': ['AccumulatedOtherComprehensiveIncomeLossNetOfTax'],
            'AccumDeficit': ['RetainedEarningsAccumulatedDeficit'],
            'NoncontrollingInterest': ['MinorityInterest', 'NoncontrollingInterest', 'NonredeemableNoncontrollingInterest'],
        }

        bs_data = {k: {y: 0.0 for y in target_years} for k in bs_keys}
        raw_bs_success = False

        # Raw CF fallback containers across all 5 years
        raw_cf_da = {y: 0.0 for y in target_years}
        raw_cf_sbc = {y: 0.0 for y in target_years}
        raw_cf_capex = {y: 0.0 for y in target_years}

        try:
            print("Attempting to extract Balance Sheet and Cash Flow details from raw filings...")
            raw_filings_data = []
            filings_to_load_bs = filings[:min(4, len(filings))]

            for filing in filings_to_load_bs:
                xbrl = filing.xbrl()

                # Also check diluted shares from this filing if any year is still 1.0
                if any(is_data['DilutedShares'][y] == 1.0 for y in target_years):
                    f_shares = extract_shares_from_xbrl_facts(xbrl, target_years)
                    for y, s_val in f_shares.items():
                        if is_data['DilutedShares'][y] == 1.0 and s_val > 0:
                            is_data['DilutedShares'][y] = s_val

                # 1. Raw Balance Sheet
                bs_stmt_raw = None
                for stmt_dict in xbrl.statements.statements:
                    if stmt_dict.get('type') == 'BalanceSheet':
                        bs_stmt_raw = xbrl.statements.get(stmt_dict['role_name'])
                        break
                if bs_stmt_raw:
                    df_bs = bs_stmt_raw.to_dataframe().reset_index()
                    f_data = {k: {y: 0.0 for y in target_years} for k in bs_keys}
                    years_found = []
                    for y in target_years:
                        col_name = find_year_column(df_bs, y)
                        if col_name:
                            years_found.append((y, col_name))

                    for y, col_name in years_found:
                        seen_concepts = set()
                        tot_curr_assets = 0.0
                        tot_assets = 0.0
                        tot_curr_liab = 0.0
                        tot_liab = 0.0
                        tot_liab_and_eq = 0.0
                        tot_eq = 0.0

                        for idx, row in df_bs.iterrows():
                            if not is_valid_statement_row(row):
                                continue
                            concept = row.get('concept')
                            if not (pd.notna(concept) and isinstance(concept, str)):
                                continue
                            val = row.get(col_name)
                            if not pd.notna(val):
                                continue
                            if isinstance(val, pd.Series):
                                val = val.iloc[0]
                            fval = float(val)

                            clean_c = clean_concept(concept).lower()
                            if clean_c == 'assetscurrent':
                                tot_curr_assets = fval
                                continue
                            elif clean_c == 'assets':
                                tot_assets = fval
                                continue
                            elif clean_c == 'liabilitiescurrent':
                                tot_curr_liab = fval
                                continue
                            elif clean_c == 'liabilities':
                                tot_liab = fval
                                continue
                            elif clean_c == 'liabilitiesandstockholdersequity':
                                tot_liab_and_eq = fval
                                continue
                            elif clean_c in ['stockholdersequity', 'stockholdersequityincludingportionattributabletononcontrollinginterest']:
                                if tot_eq == 0.0:
                                    tot_eq = fval
                                continue

                            if concept in seen_concepts:
                                continue
                            mapped_key = map_raw_bs_concept(concept)
                            if mapped_key:
                                seen_concepts.add(concept)
                                f_data[mapped_key][y] += fval

                        # Reconcile each section against reported XBRL Balance Sheet section totals
                        curr_assets_keys = ['CashEquivalents', 'ShortTermInvestments', 'RestrictedCashCurrent', 'AccountsReceivable', 'PrepaidExpenses']
                        noncurr_assets_keys = [
                            'RestrictedCashNoncurrent', 'RestrictedInvestmentsNoncurrent', 'InvestmentsNoncurrent',
                            'EquityMethodInvestments', 'PP_E', 'LeaseAssets', 'Intangibles', 'Goodwill', 'OtherAssets'
                        ]
                        curr_liab_keys = ['AccountsPayable', 'ShortTermDebt', 'LeaseLiabilitiesCurrent', 'AccruedLiabilitiesCurrent']
                        noncurr_liab_keys = ['LongTermInsuranceReserves', 'LongTermDebt', 'LeaseLiabilitiesNoncurrent', 'OtherLiabilitiesNoncurrent']

                        if tot_curr_assets > 0:
                            ca_sum = sum(f_data[k][y] for k in curr_assets_keys)
                            f_data['PrepaidExpenses'][y] += (tot_curr_assets - ca_sum)

                        if tot_assets > 0:
                            ca_sum = sum(f_data[k][y] for k in curr_assets_keys)
                            nca_sum = sum(f_data[k][y] for k in noncurr_assets_keys)
                            f_data['OtherAssets'][y] += (tot_assets - ca_sum - nca_sum)

                        if tot_curr_liab > 0:
                            cl_sum = sum(f_data[k][y] for k in curr_liab_keys)
                            f_data['AccruedLiabilitiesCurrent'][y] += (tot_curr_liab - cl_sum)

                        target_total = tot_assets if tot_assets > 0 else tot_liab_and_eq
                        if tot_liab > 0:
                            cl_sum = sum(f_data[k][y] for k in curr_liab_keys)
                            ncl_sum = sum(f_data[k][y] for k in noncurr_liab_keys)
                            f_data['OtherLiabilitiesNoncurrent'][y] += (tot_liab - cl_sum - ncl_sum)
                        elif target_total > 0 and tot_eq != 0.0:
                            implied_liab = target_total - tot_eq - f_data['RedeemableNCI'][y]
                            cl_sum = sum(f_data[k][y] for k in curr_liab_keys)
                            ncl_sum = sum(f_data[k][y] for k in noncurr_liab_keys)
                            f_data['OtherLiabilitiesNoncurrent'][y] += (implied_liab - cl_sum - ncl_sum)

                        if target_total > 0:
                            total_liab_now = sum(f_data[k][y] for k in curr_liab_keys) + sum(f_data[k][y] for k in noncurr_liab_keys)
                            # Ensure Total Liabilities + RedeemableNCI + (APIC + AccumOCI + AccumDeficit) == Total Assets
                            f_data['APIC'][y] = target_total - total_liab_now - f_data['RedeemableNCI'][y] - f_data['AccumOCI'][y] - f_data['AccumDeficit'][y]

                    raw_filings_data.append((filing.filing_date, f_data, [y for y, _ in years_found]))

                # 2. Raw Cash Flow Statement (for D&A, SBC, and CapEX across all 5 years)
                cf_stmt_raw = None
                for stmt_dict in xbrl.statements.statements:
                    if stmt_dict.get('type') == 'CashFlowStatement':
                        cf_stmt_raw = xbrl.statements.get(stmt_dict['role_name'])
                        break
                if cf_stmt_raw:
                    df_cf_raw = cf_stmt_raw.to_dataframe().reset_index()
                    for y in target_years:
                        col_name = find_year_column(df_cf_raw, y)
                        if not col_name:
                            continue
                        da_combined = 0.0
                        da_parts = 0.0
                        seen_cf = set()
                        for idx, row in df_cf_raw.iterrows():
                            if not is_valid_statement_row(row):
                                continue
                            concept = row.get('concept')
                            if not (pd.notna(concept) and isinstance(concept, str)):
                                continue
                            if concept in seen_cf:
                                continue
                            val = row.get(col_name)
                            if not pd.notna(val):
                                continue
                            if isinstance(val, pd.Series):
                                val = val.iloc[0]
                            fval = float(val)
                            clean_c = clean_concept(concept).lower()

                            if clean_c in ['depreciationdepletionandamortization', 'depreciationandamortization', 'depreciationamortizationandaccretionnet']:
                                seen_cf.add(concept)
                                da_combined = max(da_combined, abs(fval))
                            elif clean_c in ['depreciation', 'amortizationofintangibleassets', 'amortizationofintangibles']:
                                seen_cf.add(concept)
                                da_parts += abs(fval)
                            elif clean_c in ['sharebasedcompensation', 'sharebasedpaymentarrangementnoncashexpense', 'allocatedsharebasedcompensationexpense']:
                                if raw_cf_sbc[y] == 0.0:
                                    seen_cf.add(concept)
                                    raw_cf_sbc[y] = abs(fval)
                            elif clean_c in [
                                'paymentstoacquirepropertyplantandequipment',
                                'paymentstoacquirepropertyplantandequipmentgross',
                                'paymentstoacquireproductiveassets',
                                'paymentstoacquireproductiveassetsgross',
                                'capitalexpenditures',
                                'paymentsforcapitalimprovements',
                            ]:
                                if raw_cf_capex[y] == 0.0:
                                    seen_cf.add(concept)
                                    raw_cf_capex[y] = -abs(fval)

                        da_val = da_combined if da_combined > 0 else da_parts
                        if raw_cf_da[y] == 0.0 and da_val > 0:
                            raw_cf_da[y] = da_val

            if raw_filings_data:
                raw_filings_data_sorted = sorted(raw_filings_data, key=lambda x: x[0])
                for date, f_data, years in raw_filings_data_sorted:
                    for y in years:
                        for k in bs_keys:
                            bs_data[k][y] = f_data[k][y]
                raw_bs_success = True
                print("Raw Balance Sheet extraction successful.")
        except Exception as raw_e:
            print(f"Warning: Raw Balance Sheet extraction failed: {str(raw_e)}. Falling back to standardized.")

        if not raw_bs_success:
            print("Using standardized Balance Sheet mapping...")
            for key, concepts in bs_keys.items():
                vals_dict = get_concept_values(df_bal, concepts, selected_cols_std)
                bs_data[key] = {y: vals_dict[f"FY {y}"] for y in target_years}

        cf_keys = df_cf.index.tolist()
        wc_asset_keys = []
        wc_liab_keys = []
        for key in cf_keys:
            clean_k = clean_concept(key)
            if clean_k.startswith('IncreaseDecreaseIn'):
                if any(x in clean_k for x in ['Receivable', 'Inventory', 'Prepaid', 'OperatingAssets', 'OtherOperatingAssets']):
                    wc_asset_keys.append(key)
                elif any(x in clean_k for x in ['Payable', 'Accrued', 'OperatingLiabilities', 'OtherOperatingLiabilities', 'EmployeeRelatedLiabilities']):
                    wc_liab_keys.append(key)

        wc_assets_raw = get_concept_values(df_cf, wc_asset_keys, selected_cols_std)
        wc_liabs_raw = get_concept_values(df_cf, wc_liab_keys, selected_cols_std)

        cf_data = {
            'CF_WorkingCapitalAssets': {y: wc_assets_raw[f"FY {y}"] for y in target_years},
            'CF_WorkingCapitalLiabilities': {y: wc_liabs_raw[f"FY {y}"] for y in target_years},
        }

        net_cf_operating = get_concept_values(df_cf, ['NetCashProvidedByUsedInOperatingActivities'], selected_cols_std)
        n_inc = get_concept_values(df_cf, ['NetIncomeLoss', 'ProfitLoss'], selected_cols_std)
        da_cf = get_concept_values(df_cf, ['DepreciationDepletionAndAmortization', 'DepreciationAndAmortization', 'Depreciation'], selected_cols_std)
        sbc_cf = get_concept_values(df_cf, ['ShareBasedCompensation', 'ShareBasedPaymentArrangementNoncashExpense'], selected_cols_std)

        for y in target_years:
            std_col = f"FY {y}"
            if da_cf[std_col] == 0.0 and raw_cf_da[y] > 0:
                da_cf[std_col] = raw_cf_da[y]
            if sbc_cf[std_col] == 0.0 and raw_cf_sbc[y] > 0:
                sbc_cf[std_col] = raw_cf_sbc[y]

        cf_data['CF_OtherNonCurrentAssets'] = {}
        cf_data['CF_OtherNonCurrentLiabilities'] = {}
        for y in target_years:
            std_col = f"FY {y}"
            sum_known = n_inc[std_col] + da_cf[std_col] + sbc_cf[std_col] + wc_assets_raw[std_col] + wc_liabs_raw[std_col]
            diff = net_cf_operating[std_col] - sum_known
            cf_data['CF_OtherNonCurrentAssets'][y] = diff
            cf_data['CF_OtherNonCurrentLiabilities'][y] = 0.0

        capex_raw = get_concept_values(df_cf, ['PaymentsToAcquirePropertyPlantAndEquipment', 'PaymentsToAcquirePropertyPlantAndEquipmentGross', 'PaymentsToAcquireProductiveAssets'], selected_cols_std)
        cf_data['CF_Capex'] = {}
        for y in target_years:
            std_val = -abs(capex_raw[f"FY {y}"])
            cf_data['CF_Capex'][y] = std_val if std_val != 0.0 else raw_cf_capex[y]

        proc_debt = get_concept_values(df_cf, ['ProceedsFromIssuanceOfLongTermDebt'], selected_cols_std)
        rep_debt = get_concept_values(df_cf, ['RepaymentsOfLongTermDebt'], selected_cols_std)
        cf_data['CF_LongTermDebt'] = {y: proc_debt[f"FY {y}"] - rep_debt[f"FY {y}"] for y in target_years}

        cf_data['CF_CommercialPaper'] = {y: 0.0 for y in target_years}

        rep_stock = get_concept_values(df_cf, ['PaymentsForRepurchaseOfCommonStock', 'PaymentsRelatedToTaxWithholdingForShareBasedCompensation'], selected_cols_std)
        cf_data['CF_ShareRepurchases'] = {y: -abs(rep_stock[f"FY {y}"]) for y in target_years}

        divs = get_concept_values(df_cf, ['PaymentsOfDividends', 'PaymentsOfDividendsCommonStock'], selected_cols_std)
        cf_data['CF_CommonDividends'] = {y: -abs(divs[f"FY {y}"]) for y in target_years}

        is_data['StockBasedCompensation'] = {y: sbc_cf[f"FY {y}"] for y in target_years}

        # Reconcile Depreciation & Amortization: fallback to Cash Flow statement if raw IS is 0.0
        for y in target_years:
            std_col = f"FY {y}"
            if is_data['DepreciationAmortizationIS'][y] == 0.0:
                is_data['DepreciationAmortizationIS'][y] = abs(da_cf[std_col])

        # Reconcile Income Statement expenses & calculate interest expense
        is_data['InterestExpense'] = {y: 0.0 for y in target_years}
        is_data['GeneralAdministrative'] = {y: 0.0 for y in target_years}

        for y in target_years:
            rev = abs(is_data['Revenue'][y])
            cogs = abs(is_data['COGS'][y])
            op_inc = is_data['OperatingIncomeRaw'][y]

            rd = abs(is_data['ResearchDevelopment'][y])
            sm = abs(is_data['SalesMarketing'][y])
            op_supp = abs(is_data['OperationSupport'][y])
            da = abs(is_data['DepreciationAmortizationIS'][y])

            if has_raw_is_da[y]:
                # D&A was reported as a separate operating expense on the Income Statement
                gp = rev - cogs
                target_op_exp = gp - op_inc
                ga_residual = max(0.0, target_op_exp - (rd + sm + op_supp + da))
                cogs_adj = cogs
            else:
                # D&A was pulled from Cash Flow and is embedded in COGS / OpEx on the GAAP Income Statement.
                # Deduct D&A from COGS (and any excess from G&A) so both Operating Income (Row 25) and EBITDA (Row 37) are exact.
                gp_raw = rev - cogs
                target_op_exp_ex_da = gp_raw - op_inc
                ga_ex_da = max(0.0, target_op_exp_ex_da - (rd + sm + op_supp))
                if cogs >= da:
                    cogs_adj = cogs - da
                    ga_residual = ga_ex_da
                else:
                    cogs_adj = 0.0
                    ga_residual = max(0.0, ga_ex_da - (da - cogs))

            is_data['COGS'][y] = -cogs_adj
            is_data['ResearchDevelopment'][y] = rd
            is_data['SalesMarketing'][y] = sm
            is_data['OperationSupport'][y] = op_supp
            is_data['DepreciationAmortizationIS'][y] = da
            is_data['GeneralAdministrative'][y] = ga_residual

            int_inc = abs(is_data['InterestIncomeRaw'][y])
            int_exp = abs(is_data['InterestExpenseRaw'][y])
            is_data['InterestExpense'][y] = int_inc - int_exp

    except Exception as e:
        print(f"Error fetching standardized Balance Sheet and Cash Flow: {str(e)}")
        sys.exit(1)

    # Combine data
    out_data = {}
    out_data.update(is_data)
    out_data.update(bs_data)
    out_data.update(cf_data)

    # Metadata injection
    out_data['CompanyName'] = {y: company.name for y in target_years}

    # Extract and inject latest quarterly values
    q_cash, q_debt, q_shares = get_latest_quarterly_data(company)
    out_data['LatestQuarterlyCash'] = {y: q_cash for y in target_years}
    out_data['LatestQuarterlyDebt'] = {y: q_debt for y in target_years}
    out_data['LatestQuarterlyShares'] = {y: q_shares for y in target_years}

    # Save to CSV
    csv_path = "edgar_data.csv"
    with open(csv_path, "w") as f:
        header = "Key," + ",".join([str(y) for y in target_years]) + "\n"
        f.write(header)

        for key in out_data:
            if key in ['InterestExpenseRaw', 'InterestIncomeRaw', 'GeneralAdministrativeRaw', 'OperatingIncomeRaw']:
                continue
            row_vals = []
            for y in target_years:
                val = out_data[key].get(y, 0.0)
                if key != 'DilutedEPS' and key != 'CompanyName':
                    val = val / 1000000.0
                if key == 'CompanyName':
                    row_vals.append(str(val).replace(",", ""))
                else:
                    row_vals.append(f"{val:.6f}")
            f.write(f"{key}," + ",".join(row_vals) + "\n")

    print(f"Data saved to {csv_path} successfully.")


if __name__ == "__main__":
    main()
