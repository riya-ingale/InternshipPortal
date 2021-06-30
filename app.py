from flask import Flask, render_template, url_for, request, redirect, session, flash, send_file
from flask import make_response, session, g
import pdfkit
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from io import BytesIO
from werkzeug.security import generate_password_hash, check_password_hash
import secrets
import os
# from flask_admin import Admin
# from flask_admin.contrib.sqla import ModelView
import smtplib

app = Flask(__name__) 

app.config['SECRET_KEY'] = os.urandom(24)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db= SQLAlchemy(app)
 
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = "You need to Login first"

@login_manager.user_loader
def load_user(user_id):
    return Users.query.get(int(user_id))

class Users(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    fullname = db.Column(db.String(150))
    rollno = db.Column(db.String(15), unique=True)
    password =  db.Column(db.String(80))
    mobileno = db.Column(db.String(15))
    email = db.Column(db.String(50))
    dept = db.Column(db.String(15))
    div = db.Column(db.String(15))

class Internships(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    userid= db.column(db.Integer)
    compnayname = db.column(db.Text)
    domain = db.column(db.Text)
    rating = db.column(db.Integer)
    startdate = db.column(db.Date)
    endtdate = db.column(db.Date)
    offerletter = db.Column(db.LargeBinary)
    completioncert = db.Column(db.LargeBinary)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/profile')
def profile():
    return render_template("profile.html")

@app.route('/search')
def search():
    return render_template("search.html")

@app.route('/newinternship')
def newinternship():
    return render_template("newinternship.html")


@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == "POST":
        fullname = request.form['fullname']
        rollno = request.form['rollno']
        email = request.form['email']
        mobileno = request.form['mobileno']
        dept = request.form['dept']
        div = request.form['div']
        password = request.form['password']
        hashed_password = generate_password_hash(password, method="sha256")
        cpassword = request.form['cpassword']

        user = Users.query.filter_by(rollno=rollno).first()
        if user:
            flash("Roll number Already Exists")
            return redirect("/signup")
        user = Users.query.filter_by(email=email).first()
        if user:
            flash("Email Already Registered, Try logging in")
            return redirect("/login")

        if(password == cpassword):
            new_user = Users(fullname=fullname, rollno=rollno, email=email, mobileno=mobileno, 
            dept=dept, div=div, password=hashed_password)
            db.session.add(new_user)
            db.session.commit()

            flash("Sucessfully Registered!", "success")
            return redirect('/login')
        else:
            flash("Passwords don't match", "danger")
            return redirect("/signup")

    return render_template("sign-up.html")


@app.route('/login', methods=['POST', 'GET'])
def login():
    if request.method == "POST":
        logout_user()
        rollno = request.form.get('rollno')
        password = request.form.get('password')

        user = Users.query.filter_by(rollno=rollno).first()
        if not user:
            flash("No such User found, Try Signing Up First", "warning")
            return redirect("/signup")
        if user:
            if check_password_hash(user.password, password):
                login_user(user)
                return redirect("/")
            else:
                flash("Incorrect password", "danger")
                return redirect("login")
    return render_template("login.html")


@app.route('/logout')
def logout():
    logout_user()
    flash("Successfully Logged out!")
    return redirect('/login')


if __name__ == '__main__':
    db.create_all()
    app.run(debug=True)