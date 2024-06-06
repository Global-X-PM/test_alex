import pandas as pd
import datetime as dt
import sys
sys.path.append('Z:\\IPS\\python')
import common
import bloomberg_session


if __name__ == "__main__":
    list_tickers = ["USSX","CNDX","QQQX","QQQX/U","EAFX","EAFX/U","EMMX","EMMX/U", "TTTX","AIGO","EACC","EMCC","EMML","USSL","EAFL","QQQL","PAYS","EMCL","EACL","EQCC"]

    #convert to bbg tickers
    bbg_list_ticker = [item + " CN Equity" for item in list_tickers]

    bdp = bloomberg_session.BDP_Session()
    data = bdp.bdp_request(bbg_list_ticker, ["volume","EQY_TURNOVER_REALTIME","NUM_TRADES_RT"])

    timestamp_dict = bdp.unpact_dictionary(data, "timestamp")
    volume_dict = bdp.unpact_dictionary(data, "volume")
    turnover_dict = bdp.unpact_dictionary(data, "EQY_TURNOVER_REALTIME")
    trades_count_dict = bdp.unpact_dictionary(data, "NUM_TRADES_RT")

    output = pd.DataFrame(bbg_list_ticker, columns=["ticker"])
    output["volume"] = output["ticker"].map(volume_dict)
    output["turnover today"] = output["ticker"].map(turnover_dict)
    output["# of trades"] = output["ticker"].map(trades_count_dict)
    output["timestamp"] = output["ticker"].map(timestamp_dict)
    output.sort_values(by='volume', ascending=False, inplace=True)

    output["volume"] = output["volume"].apply(lambda x: "{:,.0f}".format(x))
    output["turnover today"] = output["turnover today"].apply(lambda x: "{:,.0f}".format(x))
    output.fillna('-', inplace=True)


    css = """<style type="text/css">
    table {
        border-collapse: collapse;
        width: 80%;
    }
    th, td {
        border: 1px solid black;
        text-align: center;
        padding: 2px;
    }
    th {
        background-color: #f2f2f2;
    }
    </style>
    """


    import win32com.client as win32

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = 'hsahn@globalx.ca;asmahtin@globalx.ca'
    #mail.BCC = 'asmahtin@globalx.ca,bwok@globalx.ca'
    mail.Subject = 'Dashboard: ' + dt.datetime.now().strftime('%Y-%m-%d')
    mail.HTMLBody = css + output.to_html(index=False,classes='table table-striped')
    mail.Send()