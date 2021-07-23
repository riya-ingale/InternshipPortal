from flask import Flask, render_template, url_for, request, redirect, session, flash, send_file
from flask import make_response, session, g
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from io import BytesIO
from werkzeug.security import generate_password_hash, check_password_hash
import os
from datetime import datetime
import pdfkit
import flask_excel as excel

app = Flask(__name__)

app.config['SECRET_KEY'] = os.urandom(24)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///InternshipPortal_Database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

excel.init_excel(app)

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
    year = db.Column(db.String(15))
    linkedIn_username = db.Column(db.Text)
    role = db.Column(db.Text)


class Internships(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer)
    companyname = db.Column(db.Text)
    domain = db.Column(db.Text)
    source = db.Column(db.Text)
    rating = db.Column(db.Integer)
    skills_acquired = db.Column(db.Text)
    companyrepresentative_name = db.Column(db.Text)
    companyrepresentative_contact = db.Column(db.Text)
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
        year = request.form['year']
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
                             dept=dept, div=div, password=hashed_password, year=year)
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
@app.route('/profile/', methods=['GET', 'POST'])
@login_required
def profile(user_id):
    user = Users.query.filter_by(id=user_id).first()
    internships = Internships.query.filter_by(user_id=user_id).all()
    return render_template("profile.html", current_user=current_user, internships=internships, user=user)


@app.route('/otherprofile/<int:user_id>', methods=['GET', 'POST'])
def otherprofile(user_id):
    user = Users.query.filter_by(id=user_id).first()
    internships = Internships.query.filter_by(user_id=user_id).all()
    return render_template("otherprofile.html", current_user=current_user, internships=internships, user=user)


@app.route('/search', methods=['GET', 'POST'])
def search():
    allinternships = []
    allstudents = []
    student_data = []
    internship_data = []
    if request.method == "GET":
        allstudents = Users.query.filter(Users.id != 2).all()
        allinternships = Internships.query.all()
        return render_template("search.html", students=allstudents, internships=allinternships, s = True)

    if request.method == 'POST':
        searchname = request.form.get('searchname')
        dept = request.form.get('dept')
        div = request.form.get('div')
        year = request.form.get('year')
        startdate = request.form.get('startdate')
        print(startdate)
        if startdate:
            startdate = datetime.strptime(startdate, '%Y-%m-%d')
        enddate = request.form.get('enddate')
        print(enddate)
        if enddate:
            enddate = datetime.strptime(enddate, '%Y-%m-%d')

        search = "{0}".format(searchname)
        search = search+'%'
        print(search)

        if startdate and enddate and searchname:
            allinternships = Internships.query.filter(
                or_(Internships.companyname.like(search), Internships.domain.like(search)), Internships.startdate > startdate and Internships.enddate < enddate).all()

        elif startdate and searchname and not enddate:
            allinternships = Internships.query.filter(
                or_(Internships.companyname.like(search), Internships.domain.like(search)), Internships.startdate > startdate).all()
        elif enddate and searchname and not startdate:
            allinternships = Internships.query.filter(
                or_(Internships.companyname.like(search), Internships.domain.like(search)), Internships.endate < enddate).all()
        elif startdate and enddate and not searchname:
            allinternships = Internships.query.filter(
                Internships.startdate > startdate and Internships.enddate < enddate).all()
        elif searchname and not enddate and not startdate :
            allinternships = Internships.query.filter(
                or_(Internships.companyname.like(search), Internships.domain.like(search))).all()
        elif startdate and not searchname and not enddate:
            allinternships = Internships.query.filter(
                Internships.startdate > startdate).all()
        elif enddate and not startdate and not searchname:
            allinternships = Internships.query.filter(
                Internships.enddate < enddate).all()
        else:
            pass
        if dept and div and year:
            allstudents = Users.query.filter_by(
                dept=dept, div=div, year=year).all()
        elif dept and div and not year:
            allstudents = Users.query.filter_by(dept=dept, div=div).all()
        elif div and year and not dept:
            allstudents = Users.query.filter_by(div=div, year=year).all()
        elif dept and year and not div:
            allstudents = Users.query.filter_by(dept=dept, year=year).all()
        elif dept and not div and not year:
            allstudents = Users.query.filter_by(dept=dept).all()
        elif div and not dept and not year:
            allstudents = Users.query.filter_by(div=div).all()
        elif year and not div and not dept:
            allstudents = Users.query.filter_by(year=year).all()
        else:
            pass
        if allinternships and not allstudents:
            for internship in allinternships:
                student = Users.query.filter_by(id=internship.user_id).first()
                student_data.append(student)
            return render_template("search.html", students=student_data, internships=allinternships, s = True)
        elif allstudents and not allinternships:
            for student in allstudents:
                student = Internships.query.filter_by(
                    user_id=student.id).first()
                internship_data.append(student)
            return render_template("search.html", students=allstudents, internships=internship_data, s = True)
        elif allinternships and allstudents:
            return render_template("search.html", students=allstudents, internships=allinternships, s = True)
        else:
            return render_template("search.html",s = False)


@app.route('/newinternship', methods=['GET', 'POST'])
@login_required
def newinternship():
    if request.method == "POST":
        companyname = request.form.get('companyname')
        domain = request.form.get('domain')
        source = request.form.get('source')
        rating = request.form.get('rating')
        skills_acquired = request.form.get('skills_acquired')
        companyrepresentative_name = request.form.get(
            'companyrepresentative_name')
        companyrepresentative_contact = request.form.get(
            'companyrepresentative_contact')
        startdate = request.form.get('startdate')
        startdate = datetime.strptime(startdate, '%Y-%m-%d')
        enddate = request.form.get('enddate')
        enddate = datetime.strptime(enddate, '%Y-%m-%d')
        offerletter = request.files['offerletter']
        completioncert = request.files['completioncert']
        if len(offerletter.filename) > 0:
            offerletter_filename = offerletter.filename
            offerletter = offerletter.read()
        if len(completioncert.filename) > 0:
            completioncert_filename = completioncert.filename
            completioncert = completioncert.read()
        new_internship = Internships(user_id=current_user.id, companyname=companyname, domain=domain, companyrepresentative_name=companyrepresentative_name, companyrepresentative_contact=companyrepresentative_contact, source=source, rating=rating, skills_acquired=skills_acquired, startdate=startdate, enddate=enddate,
                                     offerletter=offerletter, offerletter_filename=offerletter_filename, completioncert=completioncert, completioncert_filename=completioncert_filename)
        db.session.add(new_internship)
        db.session.commit()
        flash("Record Added!")
        return redirect('/newinternship')
    return render_template("newinternship.html")


@app.route('/updateinternship/<int:id>', methods = ['GET','POST'])
@login_required
def updateinternship(id):
    internship = Internships.query.filter_by(id = id).first()
    print(internship.companyname)
    if request.method == 'POST':
        internship.companyname = request.form.get('companyname')
        internship.domain = request.form.get('domain')
        internship.source = request.form.get('source')
        internship.rating = request.form.get('rating')
        internship.skills_acquired = request.form.get('skills_acquired')
        internship.companyrepresentative_name = request.form.get(
            'companyrepresentative_name')
        internship.companyrepresentative_contact = request.form.get(
            'companyrepresentative_contact')
        startdate = request.form.get('startdate')
        internship.startdate = datetime.strptime(startdate, '%Y-%m-%d')
        enddate = request.form.get('enddate')
        internship.enddate = datetime.strptime(enddate, '%Y-%m-%d')
        offerletter = request.files['offerletter']
        completioncert = request.files['completioncert']
        if len(offerletter.filename) > 0:
            internship.offerletter_filename = offerletter.filename
            internship.offerletter = offerletter.read()
        if len(completioncert.filename) > 0:
            internship.completioncert_filename = completioncert.filename
            internship.completioncert = completioncert.read()
        db.session.commit()
        return redirect(f'/profile/{current_user.id}')
    return render_template('updateinternship.html', internship = internship)


@app.route('/downloadcompletioncert/<int:internship_id>', methods=['GET', 'POST'])
def downloadcompletioncert(internship_id):
    internship = Internships.query.filter_by(id=internship_id).first()
    if internship.completioncert:
        file_data = internship.completioncert
        return send_file(BytesIO(file_data), attachment_filename=internship.companyname + "Completioncert.pdf", as_attachment=True)
    else:
        flash("No file Exists")
        return redirect(f'/profile/{current_user.id}')


@app.route('/student_record/<int:user_id>', methods=['POST', 'GET'])
def student_record(user_id):
    user = Users.query.get_or_404(user_id)
    internships = Internships.query.filter_by(user_id=user_id).all()
    rendered = render_template(
        'studentrecorddownload.html', student=user, internships=internships)
    pdf = pdfkit.from_string(rendered, False)
    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename = student.rollno +_Details.pdf'

    return response


@app.route('/admin/login', methods=['GET', 'POST'])
def adminlogin():
    if request.method == 'POST':
        username = request.form.get('username')
        print(username)
        password = request.form.get('password')
        print(password)
        user = Users.query.filter_by(role="admin", fullname=username).first()
        print(user.role)
        if user:
            if user.password == password:
                login_user(user)
                return redirect('/admin/dashboard')
            else:
                flash("Password Incorrect")
                return redirect('/admin/login')
        else:
            flash("No such admin exists")
            return redirect('/admin/login')
    return render_template('adminlogin.html')


@app.route('/admin/dashboard', methods=['GET', 'POST'])
def admindashboard():
    if current_user.role == "admin":
        return render_template('admindashboard.html')
    else:
        return "Page Not Found"


@app.route('/editprofile/<int:user_id>', methods=['GET', 'POST'])
def editprofile(user_id):
    user = Users.query.get_or_404(user_id)
    internships = Internships.query.filter_by(user_id=user_id).all()
    if request.method == "POST":
        user = Users.query.get_or_404(user_id)
        user.fullname = request.form.get('fullname')
        user.rollno = request.form['rollno']
        user.email = request.form['email']
        user.mobileno = request.form['mobileno']
        user.dept = request.form['dept']
        user.div = request.form['div']
        user.year = request.form['year']
        db.session.commit()
        return redirect(f'/profile/{user_id}')
    return render_template('editprofile.html', user=user, internships = internships)


@app.route('/aboutus')
def aboutus():
    return render_template('aboutus.html')


@app.route("/customexport", methods=['GET', 'POST'])
def docustomexport():
    information = request.data
    print(type(information))

    query_sets = Internships.query.all()
    print("Excel - ", query_sets)
    column_names = ['id', 'companyname', 'domain', 'source', 'rating', 'skills_acquired',
                    'companyrepresentative_name', 'companyrepresentative_contact', 'startdate', 'enddate']
    return excel.make_response_from_query_sets(query_sets, column_names, 'xlsx', file_name="sheet")


if __name__ == '__main__':
    db.create_all()
    app.run(debug=True)
