from flask import Flask, render_template, request, redirect, url_for
from DepartmentCoporate import department  # Import department blueprint
from insuranceMIS import insurance   # Import insurance blueprint
from LedgerMIS import  ledger
from CYLYMIS import cyly
from OverheadBlend import blendovr
from hotelscrap import googlescp
from MMTscrap import mmtscp
from OnlineMarket import ecomarket
from Expenseswisemvmt import expnsmvmt

app = Flask(__name__)

# Dummy user credentials for authentication
VALID_USER = 'MJIPLSFTW@12'
VALID_PASSWORD = 'A@09&34#5422'

# Route for login page
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user_id = request.form['userId']
        password = request.form['password']
        if user_id == VALID_USER and password == VALID_PASSWORD:
            return redirect(url_for('index'))  # Redirect to main interface if credentials are correct
        else:
            error = 'Invalid ID or Password!'  # Show error if credentials are incorrect
            return render_template('login.html', error=error)
    return render_template('login.html')


# Register Blueprints
app.register_blueprint(department)
app.register_blueprint(insurance)
app.register_blueprint(ledger)
app.register_blueprint(cyly)
app.register_blueprint(blendovr)
app.register_blueprint(googlescp)
app.register_blueprint(mmtscp)
app.register_blueprint(ecomarket)
app.register_blueprint(expnsmvmt)


@app.route('/main')
def index():
    return render_template('test.html')

@app.route('/processing')
def processing():
    return render_template('processing.html')

if __name__ == '__main__':
    app.run(debug=True, port=5000)

