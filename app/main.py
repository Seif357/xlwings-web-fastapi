import datetime as dt

from fastapi import Body
import yfinance as yf
import xlwings as xw

from app import app


@app.post("/hello")
def hello(data: dict = Body(...)):
    # Grab the text sent from Script Lab
    input_text = data.get("data", "")
    
    # Logic: toggle the text
    if input_text == "Hello xlwings!":
        result = "Bye xlwings!"
    else:
        result = "Hello xlwings!"

    # Return a simple JSON response that Script Lab can read
    return {"result": result}



@app.post("/interpolation")
def run_forecast(data: dict = Body(...)):
    # 1. Load the workbook state sent by the frontend
    book = xw.Book(json=data)    # 2. Get the active sheet
    sheet1 = book.sheets[0]
    
    # 3. Your exact xlwings logic
    sheet1.range("A1").value = "X"
    sheet1.range("B1").value = "Y"
    sheet1.range("A2").options(transpose=True).value = [3, 4]
    sheet1.range("B2").options(transpose=True).value = [9, 12]
    
    # Linear Interpolation at x = 3.5
# This works!
    sheet1.range("B4").value = "interpolation Equal"
    sheet1.range("A4").value = "interpolation Equal"
    sheet1.range("B5").value = "=FORECAST.LINEAR(3.5, B2:B3, A2:A3)"    
    sheet1.range("A5").value = "=FORECAST.LINEAR(3.5, B2:B3, A2:A3)"    

    # 4. Return the instructions back to the frontend
    return book.json()
@app.post("/average")
def run_forecast(data: dict = Body(...)):
    # 1. Load the workbook state sent by the frontend
    book = xw.Book(json=data)    # 2. Get the active sheet
    sheet1 = book.sheets[0]
    
    # 3. Your exact xlwings logic
    sheet1.range("C1").value = "Average"
    sheet1.range("C2").options(transpose=True).value = [10, 20, 30, 40, 50]
    
    # Linear Interpolation at x = 3.5
# This works!
    sheet1.range("C7").value = "Equal"    
    sheet1.range("C8").value = "=Average(C2:C6)"    
    # 4. Return the instructions back to the frontend
    return book.json()

@app.post("/std")
def run_forecast(data: dict = Body(...)):
    wb = xw.Book(json=data)    # 2. Get the active sheet
    sheet1=wb.sheets[0]

    sheet1.range("D1").value="Standard Deviation"
    sheet1.range("D2").options(transpose=True).value=[10, 20, 30, 40, 50]
    sheet1.range("D7").value="Equal"
    sheet1.range("D8").value = "=STDEV.S(D2:D6)"
    return wb.json()
@app.post("/Variance")
def run_forecast(data: dict = Body(...)):
    wb = xw.Book(json=data)    # 2. Get the active sheet
    sheet= wb.sheets[0]
    sheet.range("E1").value="Variance"
    sheet.range("E2").options(transpose=True).value=[10, 20, 30, 40, 50]
    sheet.range("E7").value = "Equal"    
    sheet.range("E8").value="=Var(E2:E6)"


    return wb.json()

@app.post("/Correlation")
def run_forecast(data: dict = Body(...)):
    wb = xw.Book(json=data)    # 2. Get the active sheet
    sheets=wb.sheets[0]

    sheets.range("F1").value="X"
    sheets.range("F2").options(transpose=True).value=[1, 2, 3, 4, 5]
    sheets.range("G1").value="Y"
    sheets.range("G2").options(transpose=True).value=[2, 4, 6, 8, 10]
    sheets.range("F7").value = "Correlation Equal"    
    sheets.range("G7").value = "Correlation Equal"
    sheets.range("F8").value = "=CORREL(F2:F6,G2:G6)"    
    sheets.range("G8").value = "=CORREL(F2:F6,G2:G6)"




    return wb.json()
@app.post("/Covariance")
def run_forecast(data: dict = Body(...)):
    wb = xw.Book(json=data)    # 2. Get the active sheet
    sheets= wb.sheets[0]
    sheets.range("H1").value="X"
    sheets.range("H2").options(transpose=True).value=[1, 2, 3, 4, 5]
    sheets.range("I1").value="Y"
    sheets.range("I2").options(transpose=True).value=[2, 4, 6, 8, 10]
    sheets.range("H7").value = "Covariance Equal"    
    sheets.range("I7").value = "Covariance Equal"
    sheets.range("H8").value = "=COVARIANCE.S(H2:H6,I2:I6)"    
    sheets.range("I8").value = "=COVARIANCE.S(H2:H6,I2:I6)"

    return wb.json()
@app.post("/Hypothesis")
def run_forecast(data: dict = Body(...)):
    wb = xw.Book(json=data)
    sheet=wb.sheets[0]
    sheet.range("J1").value="Observed"
    sheet.range("J2").options(transpose=True).value=[8, 12, 9, 11, 10, 10]
    sheet.range("J7").value="T-test p-value"
    sheet.range("K1").value="Expected"
    sheet.range("K2").options(transpose=True).value=[10, 10, 10, 10, 10, 10]
    sheet.range("K7").value="T-test p-value"
    sheet.range("J8").value="=TTEST(J2:J6,K2:K6,2,3)"
    sheet.range("K8").value="=TTEST(J2:J6,K2:K6,2,1)"
    sheet.range("J9").value = "Hypothesis Result"
    sheet.range("J10").value = '=IF(J8<0.05,"Reject Null Hypothesis","Fail to Reject Null Hypothesis")'
    sheet.range("K9").value = "Hypothesis Result"
    sheet.range("K10").value = '=IF(J8<0.05,"Reject Null Hypothesis","Fail to Reject Null Hypothesis")'

    return wb.json()



@app.post("/yahoo")
def yahoo_finance(data: dict = Body(...)):
    """
    This is a sample function using the yfinance package to query
    Yahoo! Finance. It writes a pandas DataFrame to Excel/Google Sheets.
    """
    book = xw.Book(json=data)

    if "yahoo" not in [sheet.name for sheet in book.sheets]:
        # Insert and prepare the sheet for first use
        sheet = book.sheets.add("yahoo")
        sheet["A1"].value = [
            "Ticker:",
            "MSFT",
            "Start:",
            dt.date.today() - dt.timedelta(days=30),
            "End:",
            dt.date.today(),
        ]
        for address in ["B1", "D1", "F1"]:
            sheet[address].color = "#D9E1F2"
        for address in ["D1", "F1"]:
            sheet[address].columns.autofit()
        sheet[
            "A3"
        ].value = "'=> Adjust the colored parameters and run the script again!"
        sheet.activate()
    else:
        # Query Yahoo! Finance
        sheet = book.sheets["yahoo"]
        target_cell = sheet["A3"]
        target_cell.expand().clear_contents()
        try:
            df = yf.download(
                sheet["B1"].value,
                start=sheet["D1"].value,
                end=sheet["F1"].value,
                progress=False,
            )
            target_cell.value = df
            target_cell.offset(row_offset=1).columns.autofit()
        except Exception as e:
            target_cell.value = repr(e)

    return book.json()


if __name__ == "__main__":
    import uvicorn
    # Changed host to 0.0.0.0 for better container compatibility
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
    
