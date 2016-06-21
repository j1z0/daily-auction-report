Note if you want to test and run locally on OSX, you need to insure you have latest version of pip prior to installing dependencies.  

Run the following:

```
pip install pip --upgrade
pip install -r requirements.txt
```


## To build the package for aws deployment

for lambda deployment you must build the deployment package on a 64-bit Amazon Linux instance.

So create an Ec2 with that instance then ssh into the ec2

then

```
sudo yum install python27-devel python27-pip gcc
sudo uum install openssl-devel libffi-devel
```

now make virtualenv
```
virtualenv ~/lambda_env
source ~/lambda_env/bin/activate
```

clone repo from gitub

```
git clone https://github.com/j1z0/daily-auction-report.git
```

install stuff

```
cd daily-auction-report/lambda
pip install pip --upgrade
pip install -r requirements.txt
```

now package the script in a zip file

```
zip -9 ~/lambda_function.zip lambda_function.py
```

and add the dependencies

```
cd $VIRTUAL_ENV/lib/python2.7/site-packages
zip -r9 ~/lambda_function.zip *
cd $VIRTUAL_ENV/lib64/python2.7/site-packages
zip -r9 ~/lambda_function.zip *
```

now upload the created package to s3
```
s3put -b daily-auction-report -a <<access_key>> -s <<secret_key>> ./lambda_function.zip
```

that will upload your deployment package to s3, then you can go to the lambda terminal and upload from s3, pasting in this url:

https://s3-ap-northeast-1.amazonaws.com/daily-auction-report/home/ec2-user/lambda_function.zip






