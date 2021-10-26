from flask import Flask, redirect, url_for, render_template, request, abort
from flask import jsonify
from json.decoder import JSONDecodeError

import requests as req

app = Flask(__name__)

@app.errorhandler(404)
def resource_not_found(e):
    return jsonify(error=str(e)), 404

@app.route("/")
def homePage():
    try:
        con = req.get("http://127.0.0.1:5000/sectors_and_companies_list")
        json_data = con.json()['data']

        if len(json_data) == 0:
            abort(404, description="No Data Yet")
            # return render_template("junnk.html", content = 'No Data Yet')
        return render_template("home.html", content = json_data)
    except req.exceptions.ConnectionError as e:
        abort(404, description="Server Connection Error")
        # return render_template("junnk.html", content = 'Server Connection Error')
    except JSONDecodeError:
        abort(404, description="Some code error since JSON Decoded Error occuring")
        # return render_template("junnk.html", content = 'Some code error since JSON Decoded Error occuring')
    except Exception as e:
        abort(404, description=e)
        # return render_template("junnk.html", content = e)


@app.route("/page=<page>")
def getPage(page):
    if page == 'home.html':
        return redirect(url_for('homePage'))
    return render_template(page)


@app.route("/sector=<SectorName>")
def getSector(SectorName):
    print('yes)')
    con = req.get(f"http://127.0.0.1:5000/sector={SectorName}")
    json_data = con.json()
    return render_template("sector.html", sector_name = SectorName, content = json_data)


@app.route("/sector=<SectorName>/company=<CompanyName>")
def getStockName(SectorName, CompanyName):
    con = req.get(f"http://127.0.0.1:5000/company_details/sector={SectorName}/company={CompanyName}")
    json_data = con.json()['data']

    if len(json_data) == 1:
        abort(404, description="File does not exist")
        # return render_template("junnk.html", content = x)

    return render_template("try_stock.html", sector_name = SectorName, company_name = CompanyName, content=json_data)

@app.route("/mcda_rankings/Rank_type=<Rank_type>")
def getMCDA(Rank_type):
    con=req.get(f"http://127.0.0.1:5000/mcda_rankings?Rank_type={Rank_type}")
    json_data=con.json()
    return render_template("top10.html",rank_type=Rank_type,content=json_data)


if __name__ == '__main__':
    app.debug=True
    app.run(port=8000)
