from flask import Flask, render_template, redirect, url_for, flash, Markup, Response, request
import os
import json
from docx_mailmerge import docxmerge
from mailmerge import MailMerge
from wtforms import TextField, SelectField, FieldList, FormField, validators
from flask.ext.wtf import Form
import sys
import numpy as np
import cx_Oracle
import pandas as pd
import datetime


SECRET_KEY = 'youll_never_guess'

mypath = "/Users/mwitt/PycharmProjects/testform/app/lettertemplates"
dt = datetime.datetime.now()
now = str(dt.year)+str(dt.month)+str(dt.day)+str(dt.hour)+str(dt.minute)+str(dt.second)
#Form fields
lettertemplate = ''
eagleaccount = ''
dict2 = ()
list2=[]
dfields=[]
Dvalues = []
Dfields = []

outpath = '/Users/mwitt/PycharmProjects/testform/app/outputletters/'
global letterout
#letterout = '_MailMergeFile'+str(now)+'.docx'
v_sql = """SELECT   a.account_id Account_Number
        ,A.ACCOUNT_BALANCE Balance_Amount
        ,OAI.ORIGINAL_CREDITOR_NAME Original_Creditor
        ,oai.original_creditor_account_nbr Original_Creditor_Account_Nbr
        ,DO.LONG_NAME Debt_Owner
        ,dpa.first_name Customer_First_Name
        ,CASE WHEN DPA.MIDDLE_NAME IS NULL THEN ' ' ELSE DPA.MIDDLE_NAME END Customer_Middle_Name
        ,dpa.last_name Customer_Last_Name
        ,dadd.address_line_1 || ' ' || dadd.address_line_2 Customer_Address
        ,DADD.CITY Customer_City
        ,dadd.state_code Customer_State
        ,DADD.POSTAL_CODE Customer_Zip
        ,extract(MONTH from PL.PLACEMENT_DATE) || '/' || extract(day from pl.placement_date) || '/' || extract(year from pl.placement_date) Placement_Date
        ,extract(MONTH from OAI.CHARGEOFF_DATE) || '/' || extract(day from OAI.CHARGEOFF_DATE) || '/' || extract(year from OAI.CHARGEOFF_DATE) Chargeoff_Date
        ,nvl2(OAI.POST_CO_INTEREST,oai.post_co_interest,0.00) Post_CO_Interest
        ,nvl2(OAI.POST_CO_FEES,oai.post_co_fees,0.00) Post_CO_Fees
FROM original_account_info oai
JOIN ACCOUNT A                  ON A.account_id             = oai.account_id
JOIN placement pl               ON pl.placement_id          = oai.placement_id
JOIN load_file lf               ON lf.load_file_id          = pl.load_file_id
JOIN debt_portfolio dp          ON dp.debt_portfolio_id     = lf.debt_portfolio_id
JOIN debt_owner DO              ON DO.debt_owner_id         = dp.debt_owner_id
JOIN debt_provider dpr          ON dpr.debt_provider_id     = dp.debt_provider_id
JOIN franchise f                ON F.FRANCHISE_ID           = A.FRANCHISE_ID
JOIN franchise_group G          ON G.FRANCHISE_GROUP_ID     = F.FRANCHISE_GROUP_ID
JOIN franchise_parent fp        ON FP.FRANCHISE_PARENT_ID   = G.FRANCHISE_PARENT_ID
join xref_debtor_account xda    on xda.account_id           = oai.account_id
left outer join xref_parent_fr_debtor xpfd on xpfd.debtor_id = xda.debtor_id
                                                  and XPFD.FRANCHISE_PARENT_ID = 1
left outer join debtor_person dpa       on DPA.DEBTOR_ID            = xda.debtor_id
left outer join debtor_address dadd on DADD.DEBTOR_ADDRESS_ID = XPFD.PREFERRED_DEBTOR_ADDRESS_ID
WHERE 1=1
and A.ACCOUNT_ID ="""

app = Flask(__name__)
app.config.from_object(__name__)

class AddressEntryForm(Form):
    Additional_Info = TextField(' ',)

class AddressesForm(Form):
    """A form for one or more manually input fields"""
    addresses = FieldList(FormField(AddressEntryForm),min_entries=1)

class SimpleForm(Form):
    accountid = TextField('Account ID')
    choicesex = []
    mychoice = ()
    files = sorted([f for f in os.listdir(mypath) if not f.startswith('.')], key=lambda f: f.lower())
    for file in files:
        mychoice = (file, file)
        choicesex.append((mychoice))
    choiceex = SelectField(u'Letter Templates', choices=choicesex)

@app.route('/', methods=['post', 'get'])
def index():
    form = SimpleForm()
    global lettertemplate,eagleaccount

    if form.validate_on_submit():
        lettertemplate = form.choiceex.data
        eagleaccount = form.accountid.data

        if len(eagleaccount)== 12: #Do we have a good account number?
            allfields = do_my_work(lettertemplate,eagleaccount) #YES
            #Check for manual fields.
            list1 = get_manual_fields(allfields)
            if len(list1) > 0: #Do we have manual fields?
                return redirect(url_for('manlfields')) #YES - We have manual fields.
            else:
                df = create_new_letter(lettertemplate,eagleaccount) #NO
                #print 'Dataframe is empty?'+str(df)
                if df:
                    flash(Markup('No Eagle account information could be found! Please check Account ID.'))  # NO - Account id is not valid.
                    return redirect(url_for('index'))
                else:
                    dt = datetime.datetime.now()
                    now = str(dt.year) + str(dt.month) + str(dt.day) + str(dt.hour) + str(dt.minute) + str(dt.second)
                    letterout = '_MailMergeFile' + str(now) + '.docx'
                    flash(Markup('Your file has been saved to: ' + outpath + eagleaccount + letterout))
                    letterout = ''
                    return redirect(url_for('index'))
        else:
            flash(Markup('Bad Eagle Account ID')) #NO - Account id is not valid.
            return render_template('index.html', form=form)
    else:
        print form.errors
        return render_template('index.html', form=form)

@app.route('/manlfields', methods=['post', 'get'])
def manlfields():
    form = AddressesForm(addresses=list2)

    if form.validate_on_submit():
        global Dvalues, Dfields, mandict

        for i in form.addresses:
            #Dvalues.append(chr(34) + str(i.data['Additional_Info']) + chr(34))
            Dvalues.append(str(i.data['Additional_Info']))

        for key in form.addresses.object_data:
            Df = key.keys()
            Df = json.dumps(Df).strip('[]"')
            Dfields.append(Df)

        mandict = get_dic_from_two_lists(Dfields,Dvalues)
        create_new_letter(lettertemplate,eagleaccount)
        dt = datetime.datetime.now()
        now = str(dt.year) + str(dt.month) + str(dt.day) + str(dt.hour) + str(dt.minute) + str(dt.second)
        letterout = '_MailMergeFile' + str(now) + '.docx'
        flash(Markup('Your file has been saved to: ' + outpath + eagleaccount + letterout))
        letterout = ''
        return redirect(url_for('index'))
    else:
        print form.errors
        return render_template('manflds.html', form=form)

#Create the new letter.
def create_new_letter(lttrtmplt,eagleacct):

    # establish DB connection to prod
    dsn_prod = cx_Oracle.makedsn('<server>', '<port>', service_name='<service_name>')
    conn_prod = cx_Oracle.connect(user='<username>', password='<password>', dsn=dsn_prod)
    curr_prod = conn_prod.cursor()
    df_sql = pd.read_sql(v_sql + eagleacct, conn_prod)
    global DBlist1,DBlist2,DBdict
    DBlist1 = []
    DBlist2 = []

    #Did the query return any values?
    if df_sql.empty:
        conn_prod.close()
        return df_sql.empty #If so, return true
    else:

        #Loop through all fields and append keys, values
        #Also to find the ONE field that is too long
        for name, values in df_sql.iteritems():

            if name != 'ORIGINAL_CREDITOR_ACCOUNT_NBR':
                DBlist1.append('{name}'.format(name=name))
                DBlist2.append('{value}'.format(value=values[0]))

            elif name == 'ORIGINAL_CREDITOR_ACCOUNT_NBR': #Everyone prefers to see the word NUMBER spelled out.

                DBlist1.append('ORIGINAL_CREDITOR_ACCOUNT_NUMBER')
                DBlist2.append('{value}'.format(value=values[0]))

        conn_prod.close()

        #Append the Manual Fields keys, values if they exist
        if len(Dfields) > 0:
            for k in Dfields:
                DBlist1.append(k)

            for v in Dvalues:
                DBlist2.append(v)

        # Combine keys with values and put into a dictionary
        # and send it to the docxmerge function with Letter
        # Template and Output file.
        dt = datetime.datetime.now()
        now = str(dt.year) + str(dt.month) + str(dt.day) + str(dt.hour) + str(dt.minute) + str(dt.second)
        letterout = '_MailMergeFile' + str(now) + '.docx'
        DBdict = get_dic_from_two_lists(DBlist1, DBlist2)
        docxmerge(mypath + '/' + lttrtmplt, DBdict, outpath + eagleacct + letterout)

#Get all the mailmerge fields in a chosen mailmerge letter template docx.
def do_my_work(lettertemplatefile,eagleacctid):
    document = MailMerge(mypath+'/'+lettertemplatefile)
    dbmf = get_mergefields(document)
    return dbmf

#Get only the manually entered fields from all the merge fields in a mailmerge letter template docx.
#They are identified by the leading 'D_'.
def get_manual_fields(curmergefields):
    global list2
    list2 = []
    dict3 = ()
    for i in curmergefields:
        if i[0:2] == 'D_':
            dict3 = {i:i}
            list2.append(dict3)
    return list2

#Get all merge fields from supplied mailmerge letter template docx.
def get_mergefields(doc):
    dgmf = doc.get_merge_fields()
    return dgmf

# Combine two lists into dictionary with keys and values
def get_dic_from_two_lists(keys, values):
    return {keys[i]: values[i] for i in range(len(keys))}

if __name__ == '__main__':
   app.run(debug = True)
