# |-----------------------------------------------------------------------------
# |            This source code is provided under the Apache 2.0 license      --
# |  and is provided AS IS with no warranty or guarantee of fit for purpose.  --
# |                See the project's LICENSE.md for details.                  --
# |           Copyright Refinitiv 2020. All rights reserved.                  --
# |-----------------------------------------------------------------------------

# |-----------------------------------------------------------------------------
# | Please see more detail regarding this application in README.md file       --
# |-----------------------------------------------------------------------------


import xlwings as xw
import configparser as cp
import refinitiv.dataplatform as rdp

from refinitiv.dataplatform.content import ipa
from refinitiv.dataplatform.content.ipa import bond

"""
You should save a text file with **filename** `rdp.cfg` having the following contents:

    [rdp]
    username = YOUR_RDP_EMAIL_USERNAME
    password = YOUR_RDP_PASSWORD
    app_key = YOUR_RDP_APP_KEY
    
This file should be readily available (e.g. in the current working directory) for the next steps.
"""

session = None
wb = None

# Open RDP Platform Session
def init_session():
    cfg = cp.ConfigParser()
    """
    The cfg file must be specify in absolute path to the file, otherwise Excel file cannot read it.
    """
    cfg_location = 'C:\\drive_d\\Project\\Code\\xlwings_project\\notebook' + '\\rdp.cfg' # Change it to match your machine folder.
    cfg.read(cfg_location) 
    session = rdp.open_platform_session(
        cfg['rdp']['app_key'], 
        rdp.GrantPassword(
            username = cfg['rdp']['username'], 
            password = cfg['rdp']['password']
        )
    )
    #print(session.get_open_state())
    return session

"""
## Bond Pricing

The get_bond_analytics function computes bond analytics (yield, sensitivities, spreads) based on the latest available market data or using end of day data.

The list of available fields can be found on the Refinitiv Developer Community https://developers.refinitiv.com/en/api-catalog/refinitiv-data-platform/refinitiv-data-platform-apis/documentation page

"""
def request_ipa_bond(universe, fields):
    response = None
    try:
        response = rdp.get_bond_analytics(
            universe = universe,
            calculation_params = bond.CalculationParams(
                market_data_date="2020-07-05",
                price_side = ipa.enum_types.PriceSide.BID
            ),
            fields = fields
        )
    except Exception as exp:
        print('RDP Libraries: Function Layer exception: %s' % str(exp))
    
    return response

# Close RDP Platform Session
def close_session(session):
    rdp.close_session()
    print(session.get_open_state())

"""
The python application must have a method named "main" as a main point to run application by the macro-enabled Excel file.
The python application and Excel files name must be identical.

"""
def main():
    wb = xw.Book.caller()
    ipa_sheet = wb.sheets[0]
    ipa_sheet.name = 'IPA Bond Sheet'
    session = init_session()
    universe = ["US3MT=RR","US6MT=RR","US1YT=RR", "US2YT=RR", "US3YT=RR", "US5YT=RR", "US7YT=RR", "US10YT=RR"]
    fields = ['InstrumentDescription','MarketDataDate','Price','YieldPercent','ZSpreadBp']
    df_response = request_ipa_bond(universe, fields)
    if df_response is not None:
        df_response.insert(0, 'Item', universe, True)
        df_response.set_index('Item', inplace = True)
        #print(df_response)
        ipa_sheet['A8'].value = df_response
    
    close_session(session)
