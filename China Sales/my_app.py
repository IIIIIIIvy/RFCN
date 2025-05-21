import os
import shutil
import data_serialization
import requests
import pandas as pd
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import and_,delete,update
from flask import render_template,url_for,redirect,request,jsonify
from flask import flash
from datetime import datetime
from werkzeug.utils import secure_filename
from werkzeug.datastructures import FileStorage

USER_NAME='postgres'
PASS_WD='Radioflyer1'

app=Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI']='postgresql+psycopg2://'+USER_NAME+':'+PASS_WD+'@localhost/CS'
app.config['SQLALCHEMY_ECHO']=True
#app.config['SECRET_KEY']='ABCDEFG!'
db=SQLAlchemy(app)

basedir = os.path.abspath(os.path.dirname(__file__))

class AccountInfo(db.Model):
    __tablename__ = 'account_info'

    email = db.Column(db.Text, nullable=False,primary_key=True)
    name = db.Column(db.Text,nullable=False)
    passwd = db.Column(db.Text,nullable=False)
    operate_time = db.Column(db.DateTime, server_default=db.FetchedValue(),nullable=False)



class AfterSalesRecord(db.Model):
    __tablename__ = 'after_sales_records'

    rma_ = db.Column('rma#', db.Text, primary_key=True)
    contact_date = db.Column(db.Date)
    purchase_date = db.Column(db.Date)
    contact_id = db.Column(db.Text)
    source_of_purchase = db.Column(db.Text)
    item_ = db.Column('item#', db.Text)
    defect_unit = db.Column(db.Double(53))
    original_address = db.Column(db.Text)
    defect_description = db.Column(db.Text)
    action_to_be_taken = db.Column(db.Text)
    parts_no = db.Column(db.Text)
    tracking_ = db.Column('tracking#', db.Text)
    courier_ = db.Column(db.Text)
    complaint_category_class_i = db.Column(db.Text)
    complaint_category_class_ii = db.Column(db.Text)
    factory = db.Column(db.Text)
    name = db.Column(db.Text)
    number = db.Column(db.Text)
    address = db.Column(db.Text)
    update_time = db.Column(db.DateTime, server_default=db.FetchedValue())
    is_del = db.Column(db.Boolean, nullable=False, server_default=db.FetchedValue(), info='whether is deleted, 1(yes)/0(no)')
    pic_name = db.Column(db.Text)
    video_name = db.Column(db.Text)



@app.route('/all',methods=['GET','POST'])
def search_all():
    records=AfterSalesRecord.query.filter(AfterSalesRecord.is_del==False)\
				.order_by(AfterSalesRecord.update_time.desc(),AfterSalesRecord.rma_.desc())\
				.all()
    print("records.len:",len(records))
    print("records.type:",type(records))
   
    return render_template('demo4.html',
                           records=records
			)

@app.route('/get-data',methods=['GET'])
def get_data():
    # 从数据库获取数据
    records=AfterSalesRecord.query.filter(AfterSalesRecord.is_del==False)\
				.order_by(AfterSalesRecord.update_time.desc(),AfterSalesRecord.rma_.desc())\
				.all()
    data = [{'rma_': r.rma_ ,
              'contact_date': datetime.strftime(r.contact_date,'%Y-%m-%d'), 
              'purchase_date': '-' if r.purchase_date is None else datetime.strftime(r.purchase_date,'%Y-%m-%d'),
              'contact_id': r.contact_id or '-',
              'source_of_purchase': r.source_of_purchase or '-', 
              'factory': r.factory or '-', 
              'item_': r.item_ or '-',
              'complaint_category_class_i': r.complaint_category_class_i or '-', 
              'complaint_category_class_ii': r.complaint_category_class_ii or '-',
              'defect_description':   r.defect_description or '-', 
              'defect_unit': r.defect_unit or '-',
              'name': r.name or '-',
              'number': r.number or '-',
              'address': r.address or '-', 
              'parts_no': r.parts_no or '-',
              'action_to_be_taken': r.action_to_be_taken or '-',
              'courier_': r.courier_ or '-', 
              'tracking_': r.tracking_ or '-',
              'pic_name': r.pic_name or '',
              'video_name': r.video_name or '',
              } for r in records]

    dateStart = request.args.get('dateStart')
    dateEnd = request.args.get('dateEnd')
    
    if (len(dateStart)!=0) & (len(dateEnd)!=0):
        dateStart = datetime.strptime(dateStart,"%Y-%m-%d")
        dateEnd = datetime.strptime(dateEnd,"%Y-%m-%d")
        
        filtered_data = [item for item in data if datetime.strptime(item['contact_date'],"%Y-%m-%d")>=dateStart and datetime.strptime(item['contact_date'],"%Y-%m-%d")<=dateEnd]
        data = filtered_data
    return jsonify(rows=data)


@app.route('/get_latest_rma',methods=['GET'])
def get_latest_rma():
    db.session.commit()
    with db.session.begin():
        rma_latest = int(AfterSalesRecord.query.populate_existing().with_for_update(read=False,of=AfterSalesRecord).order_by(AfterSalesRecord.rma_.desc()).first().rma_)+1
        record = AfterSalesRecord(
                rma_=rma_latest, 
                contact_date=datetime.now(), 
                purchase_date=None, 
                contact_id=None, 
                source_of_purchase=None, 
                factory=None, 
                item_=None, 
                complaint_category_class_i=None, 
                complaint_category_class_ii=None, 
                defect_description=None, 
                defect_unit=None, 
                name=None, 
                number=None, 
                address=None, 
                parts_no=None, 
                action_to_be_taken=None, 
                courier_=None, 
                tracking_=None)
        db.session.add(record)
    db.session.commit()
    return jsonify({'rma':rma_latest})

@app.route('/get_latest_unique_factory',methods=['GET'])
def get_latest_unique_factory():
    db.session.commit()
    with db.session.begin():
        unique_factory = db.session.query(AfterSalesRecord.factory).distinct().populate_existing().all()
    db.session.commit()

    unique_factory=[factory[0] for factory in unique_factory]
    # unique_factory=list(filter(lambda x:x is not None,unique_factory))
    print('unique_factory:',unique_factory)
    return jsonify({'unique_factory':unique_factory})

@app.route('/get_latest_unique_source',methods=['GET'])
def get_latest_unique_source():
    db.session.commit()
    with db.session.begin():
        unique_source = db.session.query(AfterSalesRecord.source_of_purchase).distinct().populate_existing().all()
    db.session.commit()

    unique_source=[source[0] for source in unique_source]
    # unique_source=list(filter(lambda x:x is not None,unique_source))
    print('unique_source:',unique_source)
    return jsonify({'unique_source':unique_source})


@app.route('/add_cancel', methods=['POST'])
def add_cancel():
    print(request.get_json())
    rma_ = request.get_json()
    with db.session.begin():
        db.session.query(AfterSalesRecord).filter(AfterSalesRecord.rma_ == str(rma_)).populate_existing().with_for_update(read=False,of=AfterSalesRecord).delete(synchronize_session='fetch')
    db.session.commit()
    return jsonify({'message': 'Data deleted successfully'})


# 路由：接收数据并存入数据库
@app.route('/delete', methods=['POST'])
def delete_data():
    data = request.get_json()
    print(data)
    deleted_failed_list=[]
    
    with db.session.begin():
        for rma in data:
            record = AfterSalesRecord.query.filter(AfterSalesRecord.rma_==rma,~AfterSalesRecord.is_del).populate_existing().with_for_update(read=False,of=AfterSalesRecord).first()
            if record:
                record.is_del = True
                record.update_time = datetime.now()
                db.session.merge(record)
            else:
                deleted_failed_list.append(rma)
                print('Data deleted failed on ',rma)
    db.session.commit()

    return_message=''
    if len(deleted_failed_list)==0:
        return_message='Data deleted successfully.'
    elif len(deleted_failed_list)==len(data):
        return_message='Deleted failed on all records affected: ' + ','.join(deleted_failed_list)+'\nPlease check if anyone else deleted these records right before you.'
    else:
        return_message='Success: '+','.join(list(set(data).difference(set(deleted_failed_list))))+'\n'+'Fail: '+ ','.join(deleted_failed_list)+'\nPlease check if anyone else deleted these records right before you.'
    return jsonify({'message': return_message})

@app.route('/add', methods=['POST'])
def add_data():
    data = request.get_json()

    print('--------------------------------------------add--------------------------------------------')
    print(data) 
    for key,value in data.items():
        if len(value)==0:
            data[key]=None
    
    with db.session.begin():
        target = AfterSalesRecord.query.filter_by(rma_=data.get('rma')).populate_existing().with_for_update(read=False,of=AfterSalesRecord).first()
        target.contact_date=data.get('contact_date')
        target.purchase_date=data.get('purchase_date'), 
        target.contact_id=data.get('contact_id'), 
        target.source_of_purchase=data.get('source_of_purchase'), 
        target.factory=data.get('factory').strip().capitalize() if data.get('factory') else None, 
        target.item_=data.get('item').upper(), 
        target.complaint_category_class_i=data.get('complaint_category_class_i'), 
        target.complaint_category_class_ii=data.get('complaint_category_class_ii').strip().capitalize() if data.get('complaint_category_class_ii') else None,
        target.defect_description=data.get('defect_description').strip().capitalize() if data.get('defect_description') else None, 
        target.defect_unit=data.get('defect_unit'), 
        target.name=data.get('name'), 
        target.number=data.get('number'), 
        target.address=data.get('address'), 
        target.parts_no=data.get('parts_no'), 
        target.action_to_be_taken=data.get('action_to_be_taken'), 
        target.courier_=data.get('courier'), 
        target.tracking_=data.get('tracking')
        db.session.merge(target)
    # record = AfterSalesRecord(
    #         rma_=data.get('rma'), 
    #         contact_date=data.get('contact_date'), 
    #         purchase_date=data.get('purchase_date'), 
    #         contact_id=data.get('contact_id'), 
    #         source_of_purchase=data.get('source_of_purchase'), 
    #         factory=data.get('factory'), 
    #         item_=data.get('item'), 
    #         complaint_category_class_i=data.get('complaint_category_class_i'), 
    #         complaint_category_class_ii=data.get('complaint_category_class_ii'), 
    #         defect_description=data.get('defect_description'), 
    #         defect_unit=data.get('defect_unit'), 
    #         name=data.get('name'), 
    #         number=data.get('number'), 
    #         address=data.get('address'), 
    #         parts_no=data.get('parts_no'), 
    #         action_to_be_taken=data.get('action_to_be_taken'), 
    #         courier_=data.get('courier'), 
    #         tracking_=data.get('tracking'))
    # db.session.add(record)
    db.session.commit()
    print('add succeed')
    return jsonify({'message': 'Data added successfully.'})
    

@app.route('/add_pic', methods=['POST'])
def add_pic():
    print("------------------------------add_pic------------------------------")
    print(request.files)
    print(request.form)
    f_obj1=request.files[request.form.get('key').split('; ')[0]]
    if f_obj1 is None:
        return jsonify({'message':'Can not find the picture. Please re-load.'})
    else:
        pic_name_in_database = []
        video_name_in_database = []
        folder_name = basedir+"\\images\\"+request.form.get("media_id")+"\\"
        if request.form.get('requestFrom')=='edit':
            # 删掉所有文件包括文件夹
            shutil.rmtree(folder_name,ignore_errors=True)
        
        # if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        for key in request.form.get('key').split('; '):
            f_obj = request.files[key]
            f_name = secure_filename(f_obj.filename)
            f_name_suffix = f_name.split('.')[1]
            f_obj.filename = request.form.get("media_id")+"."+f_name_suffix

            file_name_this = request.form.get("media_id")+"_"+key+"."+f_name_suffix
            f_obj.save(folder_name + file_name_this)

            if key.find('pic') != -1:
                pic_name_in_database.append(file_name_this)
            elif key.find('video') != -1:
                video_name_in_database.append(file_name_this)
            


        # if request.form.get('old_pic_name') is not None:
        #     old_path = basedir+"\\images\\"+request.form.get("old_pic_name")
        #     # 检查文件是否存在
        #     if os.path.exists(old_path):
        #         try:
        #             # 删除文件
        #             os.remove(old_path)
        #             print('Image deleted successfully, name:'+request.form.get('old_pic_name'))
        #         except Exception as e:
        #             # 处理删除过程中可能发生的异常
        #             print('error'+str(e))

        
        with db.session.begin():
            target = AfterSalesRecord.query.filter_by(rma_=request.form.get("media_id")).populate_existing().with_for_update(read=False,of=AfterSalesRecord).first()
            target.pic_name = ";".join(pic_name_in_database) if len(pic_name_in_database)!=0 else None
            target.video_name = ";".join(video_name_in_database) if len(video_name_in_database)!=0 else None
            db.session.merge(target)
        db.session.commit()
        return jsonify({'message':'Picture uploaded successfully.'}) 

@app.route('/upload_records', methods=['POST'])
def upload_records():
    print("------------------------------upload_records------------------------------")
    f_obj=request.files['file']
    if f_obj is None:
        return jsonify({'message':'Can not get the file. Please re-upload.'})
    else:
        data=pd.read_excel(f_obj,header=1)
        for col in [ 'tracking#', 'courier_','factory',]:
            data[col]=data[col].astype(str)
            data[col]=data[col].apply(lambda x:None if x=='nan' else x)
        data['complaint_category_class_i']=data['complaint_category_class_i'].str.title()
        data['complaint_category_class_ii']=data['complaint_category_class_ii'].str.strip().str.capitalize()
        data['defect_description']=data['defect_description'].str.strip().str.capitalize()
        data['factory']=data['factory'].str.upper()
        
        data['item#']=data['item#'].str.upper()
        data['is_del']=False
        data['pic_name']=None
        data['video_name']=None
        data['update_time']=datetime.now()
        
        rma_latest = int(AfterSalesRecord.query.populate_existing().with_for_update(read=False,of=AfterSalesRecord).order_by(AfterSalesRecord.rma_.desc()).first().rma_)+1
        rma_list=[x for x in range(rma_latest,rma_latest+len(data))]
        for i in range(0,len(rma_list)):
            data.loc[i,'rma#']=str(rma_list[i])

        data.rename(columns={'rma#':'rma_',
                             'item#':'item_',
                             'courier#':'courier_',
                             'tracking#':'tracking_',
                             },inplace=True)
        records=[AfterSalesRecord(**row) for row in data.to_dict('records')]

        db.session.add_all(records)
        db.session.commit()

        data.fillna('',inplace=True)
        data['rma_']=data['rma_'].astype(int)
        data['contact_date']=data['contact_date'].dt.strftime('%Y-%m-%d')
        data['purchase_date']=data['purchase_date'].dt.strftime('%Y-%m-%d')
        data['update_time']=data['update_time'].dt.strftime('%Y-%m-%d')
        print(len(data))
    return jsonify({'message':'Records uploaded successfully.','records':data.to_dict('records')})

@app.route('/update', methods=['POST'])
def update_data():
    data = request.get_json()
    print('--------------------------------------------update--------------------------------------------')
    print(data)
    dataId = data.get('id')
    # target=db.session.get(AfterSalesRecord,dataId)
    with db.session.begin():
        target = AfterSalesRecord.query.filter(AfterSalesRecord.rma_==data.get('rma'),~AfterSalesRecord.is_del).populate_existing().with_for_update(read=False,of=AfterSalesRecord).first()
            
        if target:
            # if data.get('rma'):
            #     if AfterSalesRecord.query.filter(AfterSalesRecord.rma_==data.get('rma')).first():
            #         return jsonify({'message': "RMA# '"+data.get('rma')+"' already exists, please re-enter.\nNotice: The latest RMA# is "\
            #                     +AfterSalesRecord.query.order_by(AfterSalesRecord.rma_.desc()).first().rma_})
                
            #     target.rma_= data.get('rma')
            if data.get('contact_date'):
                target.contact_date= data.get('contact_date')
            if data.get('purchase_date'):
                target.purchase_date= data.get('purchase_date')
            if data.get('contact_id'):
                target.contact_id= data.get('contact_id')
            if data.get('source_of_purchase'):
                target.source_of_purchase= data.get('source_of_purchase')
            if data.get('factory'):
                target.factory= data.get('factory')
            if data.get('item'):
                target.item_= data.get('item').upper()
            if data.get('complaint_category_class_i'):
                target.complaint_category_class_i= data.get('complaint_category_class_i')
            if data.get('complaint_category_class_ii'):
                target.complaint_category_class_ii= data.get('complaint_category_class_ii')
            if data.get('defect_description'):
                target.defect_description= data.get('defect_description')
            if data.get('defect_unit'):
                target.defect_unit= data.get('defect_unit')
            if data.get('name'):
                target.name= data.get('name')
            if data.get('number'):
                target.number= data.get('number')
            if data.get('address'):
                target.address= data.get('address')
            if data.get('parts_no'):
                target.parts_no= data.get('parts_no')
            if data.get('action_to_be_taken'):
                target.action_to_be_taken= data.get('action_to_be_taken')
            if data.get('courier'):
                target.courier_= data.get('courier')
            if data.get('tracking'):
                target.tracking_= data.get('tracking')

            target.update_time = datetime.now()
            db.session.merge(target)
        else:
            print('update failed on ',dataId)
            return jsonify({'message': 'No such RMA# in database. Maybe this record had been deleted just right now. Please contact developer for more details.'})
    db.session.commit()
    return jsonify({'message': 'Data updated successfully.'})
    

@app.route('/',methods=['GET','POST'])
def main():
    return render_template('cover.html')
    

@app.route('/verify', methods=['POST'])
def verify_account():
    print('\n'+"-------------------------main verify-----------------------------------")
    print(request.get_json())
    email = request.get_json().get('email')
    passwd=request.get_json().get('passwd')

    if AccountInfo.query.filter(and_(AccountInfo.email==email,AccountInfo.passwd==passwd)).first():
        print('login success')
        return jsonify({'message': 'Login successfully.',
                        'name':AccountInfo.query.filter(and_(AccountInfo.email==email,AccountInfo.passwd==passwd)).first().name})
    else:
        print('login failed')
        return jsonify({'message': 'User '+email+' have not signed in yet, login failed, please sign in first.'})
    

@app.route('/addAccount',methods=['POST'])
def add_account():
    print('\n'+"-------------------------main addaccount-----------------------------------")
    print(request.get_json())
    email = request.get_json().get('email')
    passwd = request.get_json().get('passwd')
    name = request.get_json().get('name')

    
    if AccountInfo.query.filter(AccountInfo.email==email).first():
        print('sign up failed')
        return jsonify({'message': 'Email '+ email +' alread exists, sign up failed.'})
    else:
        account=AccountInfo(email=email,name=name,passwd=passwd,operate_time=datetime.now())
        db.session.add(account)
        db.session.commit()
        print('sign up success')
        return jsonify({'message': 'Sign up successfully.'})
    

@app.route('/get_data_for_visualization',methods=['GET'])
def get_data_for_visualization():
    json_data = data_serialization.get_data()
    return jsonify(json_data)

@app.route('/visualization',methods=['GET'])
def visualization():
    return render_template('dashboard.html')
'''
@app.route('/update_and_delete',methods=['GET','POST'])
def update_and_delete():
    if request.method=='GET':
        b= Bag.query.all()
        return render_template('add_records.html', bags=b)
    if request.method=='POST':
        #删
        op3=request.form.get('delete')
        if op3=='Delete commit':
            tbid=request.form.get('tbid')

            target = Bag.query.get(tbid)
            if target:
                Leasing.query.filter(Leasing.bag_id==tbid).delete()
                target = Bag.query.get(tbid)
                db.session.delete(target)
                db.session.commit()
            #不存在这样的bag
            else:
                flash("No such bag, failed to delete.")
            return redirect('/update_and_delete')
        #返回
        op4=request.form.get('return')
        if op4 =='Back':
            return redirect('/')

        # 改
        op5 = request.form.get('update')
        if op5 == 'Update commit':
            oid = request.form.get('oid')
            target = Bag.query.get(oid)

            if target:
                if request.form.get('nid'):
                    target.bag_id = request.form.get('nid')

                if request.form.get('name'):
                    target.name = request.form.get('name')

                if request.form.get('color'):
                    target.color = request.form.get('color')

                if request.form.get('manufacturer'):
                    target.manufacturer = request.form.get('manufacturer')

                if request.form.get('price'):
                    target.price_per_day = request.form.get('price')

                db.session.merge(target)
                db.session.commit()
            else:
                flash("No such bag, failed to update.")

            return redirect('/update_and_delete')
'''
if __name__=='__main__':
    app.run(host="0.0.0.0",port=5001,debug=True)
    # run:服务器IP+5001
