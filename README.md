# executive-reporting
Extract daily data from RDB, prepare a report and send to executives

You'll need credentials.py to be able to connect the database of interest. Credentials.py needs to include below variables:
* db: host for database
* user_name: username for database
* user_password: password for database
* remote_port: port for database
* sender_email: email address that will send the reports
* receiver_email: email address(es) that the reports will be sent to
* passoword: password for sender email

In order to replicate the same use case, please use data_example.sql to create the database. 

During the preparation of this repo, a RDB is created using AWS and accessed through MYSQL workbench.
