from flask import Flask, render_template, url_for, request, redirect, session, flash, send_file
from flask import make_response, session, g
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from io import BytesIO
from werkzeug.security import generate_password_hash, check_password_hash
import os

app = Flask(__name__)

app.config['SECRET_KEY'] = os.urandom(24)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///InternshipPortal_Database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

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
    password = db.Column(db.String(80))
    mobileno = db.Column(db.String(15))
    email = db.Column(db.String(50))
    dept = db.Column(db.String(15))
    div = db.Column(db.String(15))


class Internships(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer)
    companyname = db.Column(db.Text)
    domain = db.Column(db.Text)
    source = db.Column(db.Text)
    rating = db.Column(db.Integer)
    startdate = db.Column(db.Date)
    enddate = db.Column(db.Date)
    offerletter = db.Column(db.LargeBinary)
    offerletter_filename = db.Column(db.Text)
    completioncert = db.Column(db.LargeBinary)
    completioncert_filename = db.Column(db.Text)


@app.route('/')
def index():
    return render_template("index.html")


@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == "POST":
        fullname = request.form.get('fullname')
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
            flash("Email Already Registered")
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

    return render_template("login.html")


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
    return redirect('/')


@app.route('/profile/<int:user_id>', methods=['GET', 'POST'])
@login_required
def profile(user_id):
    user_id = current_user.id
    internships = Internships.query.filter_by(user_id=user_id).all()
    return render_template("profile.html", current_user=current_user, internships=internships)


@app.route('/search', methods=['GET', 'POST'])
def search():
    if request.method == "GET":
        allstudents = Users.query.all()
        return render_template("search.html", students=allstudents)


@app.route('/newinternship', methods=['GET', 'POST'])
@login_required
def newinternship():
    if request.method == "POST":
        companyname = request.form.get('companyname')
        domain = request.form.get('domain')
        source = request.form.get('source')
        rating = request.form.get('rating')
        startdate = request.form.get('startdate')
        enddate = request.form.get('enddate')
        offerletter = request.files.get('offerletter')
        completioncert = request.files.get('completioncert')
        if offerletter:
            offerletter_filename = offerletter.filename
            offerletter = offerletter.read()
        if completioncert:
            completioncert_filename = offerletter.filename
            completioncert = offerletter.read()
        new_internship = Internships(user_id=current_user.id, companyname=companyname, domain=domain,     source=source, rating=rating, startdate=startdate, enddate=enddate,
                                     offerletter=offerletter, offerletter_filename=offerletter_filename, completioncert=completioncert, completioncert_filename=completioncert_filename)
        db.session.add(new_internship)
        db.session.commit()
        flash("Record Added!")
        return redirect('/newinternship.html')
    return render_template("newinternship.html")

@app.route('/downloadcompletioncert/<int:user_id>', methods=['GET', 'POST'])
def downloadmarksheet12(user_id):
    user = Users.query.filter_by(id = user_id).first()
    if user.completioncert:
        file_data = user.completioncert
        return send_file(BytesIO(file_data), attachment_filename=user.rollno + user.companyname + "Completioncert.pdf", as_attachment=True)
    else:
        flash("No file Exists")
        return redirect(f'/profile/{user_id}')



@app.route('/admin/login')
def adminlogin():
    return render_template('adminlogin.html')


@app.route('/editprofile')
def editprofile():
    return render_template('editprofile.html')


if __name__ == '__main__':
    db.create_all()
    app.run(debug=True)
