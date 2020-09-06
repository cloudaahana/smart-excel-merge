from __future__ import absolute_import
import sys
import os

from google.cloud import bigquery

try:
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "key.json"
except:
    print("Please check if \"key.json\" file exists.")

projectid = 'rental-agency-288514'
filename = 'bigquery-data.csv'	# pass csv filename as commandline argument
dataset_id = 'PmData'
table_id = 'PM_DATA'
tablefullname = projectid+"."+dataset_id+"."+table_id

print('Connecting to BigQuery Database.')
client = bigquery.Client(project=projectid)

dataset_ref = client.dataset(dataset_id)
table_ref = dataset_ref.table(table_id)

job_config = bigquery.LoadJobConfig()
job_config.source_format = bigquery.SourceFormat.CSV
job_config.skip_leading_rows = 1
job_config.autodetect = True

job_config.schema = [
    bigquery.SchemaField('ACCOUNT_NAME','STRING'),
    bigquery.SchemaField('ACCOUNT_OWNER','STRING'),
    bigquery.SchemaField('BUSINESS_PHONE','STRING'),
    bigquery.SchemaField('COMPANY','STRING'),
    bigquery.SchemaField('COMPANY_TYPE','STRING'),
    bigquery.SchemaField('CREATED','STRING'),
    bigquery.SchemaField('UPDATED','STRING'),
    bigquery.SchemaField('DESCRIPTION','STRING'),
    bigquery.SchemaField('DOOR_COUNT','STRING'),
    bigquery.SchemaField('EMAIL','STRING'),
    bigquery.SchemaField('FIRST_NAME','STRING'),
    bigquery.SchemaField('LAST_NAME','STRING'),
    bigquery.SchemaField('GROWTH_PLAN','STRING'),
    bigquery.SchemaField('GROWTH_TARGET_2020','STRING'),
    bigquery.SchemaField('LEAD_OWNER','STRING'),
    bigquery.SchemaField('LEAD_STATUS','STRING'),
    bigquery.SchemaField('MAILING_CITY','STRING'),
    bigquery.SchemaField('MAILING_COUNTRY','STRING'),
    bigquery.SchemaField('MAILING_STATE','STRING'),
    bigquery.SchemaField('MAILING_STREET','STRING'),
    bigquery.SchemaField('MAILING_ZIP','STRING'),
    bigquery.SchemaField('MOBILE_PHONE','STRING'),
    bigquery.SchemaField('PARTNER','STRING'),
    bigquery.SchemaField('SOURCE','STRING'),
    bigquery.SchemaField('OTHER_CITY','STRING'),
    bigquery.SchemaField('OTHER_STATE','STRING'),
    bigquery.SchemaField('OTHER_COUNTRY','STRING'),
    bigquery.SchemaField('OTHER_ZIP','STRING'),
    bigquery.SchemaField('OTHER_STREET','STRING'),
    bigquery.SchemaField('WEBSITE_1','STRING'),
    bigquery.SchemaField('WEBSITE_2','STRING'),
    bigquery.SchemaField('SURVEY_CONDUCTED','STRING'),
    bigquery.SchemaField('DEAL_NAME','STRING'),
    bigquery.SchemaField('CLOSING_DATE','STRING'),
    bigquery.SchemaField('CONTACT_TYPE','STRING'),
    bigquery.SchemaField('CONTACT_NAME','STRING'),
    bigquery.SchemaField('BDMS','STRING'),
]

print('Creating Upload Job')
with open(filename, 'rb') as source_file:
    job = client.load_table_from_file(
        source_file,
        table_ref,
        location='US',  # Must match the destination dataset location.
        job_config=job_config)  # API request

print('Uploading Data into BigQuery Table')

job.result()  # Waits for table load to complete.

print('Loaded {} rows into {}:{}.'.format(job.output_rows, dataset_id, table_id))