import boto3
from openpyxl import Workbook

def assume_role_and_get_ec2_inventory(role_arn, session_name, region,aws_account):
    sts_client = boto3.client('sts')
    
    # Assume the role
    assumed_role = sts_client.assume_role(
        RoleArn=role_arn,
        RoleSessionName=session_name
    )

    # Create a new session using the assumed role credentials
    session = boto3.Session(
        aws_access_key_id=assumed_role['Credentials']['AccessKeyId'],
        aws_secret_access_key=assumed_role['Credentials']['SecretAccessKey'],
        aws_session_token=assumed_role['Credentials']['SessionToken'],
        region_name=region
    )

    print(f"RoleARN: {role_arn}")
                

    # Get EC2 instances details for the account
    ec2 = session.client('ec2')
    #ec2 = boto3.client('ec2')

    instances = ec2.describe_instances()

    ec2_inventory = []

    for reservation in instances['Reservations']:
        for instance in reservation['Instances']:
            # Disk details
            block_device_mappings = instance['BlockDeviceMappings']
            disk_details = []

            # tags
            tagstr = ''
            productvalue = ''
            namevalue = ''
            tags = instance['Tags']
            for tag in tags:
                key = tag['Key']
                value = tag['Value']
                print(f"Key: {key}, Value: {value}")
                if key == 'product':
                    productvalue = value
                if key == 'Name' or key == 'name':
                    namevalue = value
                
                tagstr = tagstr + key + " : " + value + ";"
                

            for mapping in block_device_mappings:
                device_name = mapping['DeviceName']
                volume_id = mapping['Ebs']['VolumeId']
                # Get EBS volume information
                ebs_response = ec2.describe_volumes(VolumeIds=[volume_id])
                ebs_volume = ebs_response['Volumes'][0]

                # Extract disk capacity
                disk_capacity = ebs_volume['Size']

                disk_details.append({'DeviceName': device_name, 'VolumeId': volume_id,'DiskCapacityGB': disk_capacity})

            disk_details_string = " "
            for disk_details_info in disk_details:
                disk_details_string = disk_details_info['VolumeId'] + "-" + str(disk_details_info['DiskCapacityGB']) + "GB " + " ; " + disk_details_string 
        
            instance_info = {
                'Account': aws_account,
                'Name': namevalue,
                'Region': region,
                'InstanceId': instance['InstanceId'],
                'InstanceType': instance['InstanceType'],
                'PrivateIpAddress': instance.get('PrivateIpAddress', 'N/A'),
                'PublicIpAddress': instance.get('PublicIpAddress', 'N/A'),
                # CPU details
                'cpu_details':instance['CpuOptions']['CoreCount'],

                # Memory details
                #'memory_details':instance['Memory']['SizeInMiB'],
                'memory_details':'xx gb',


                # Memory details
                'disk_details':disk_details_string,


                'State': instance['State']['Name'],
                'LaunchTime': str(instance['LaunchTime']),
                'ProductTag':productvalue,
                'Tags':tagstr
            }
            ec2_inventory.append(instance_info)

    return ec2_inventory

def export_to_excel(ec2_inventory, excel_file,wb,ws):
    #wb = Workbook()
    #ws = wb.active

    # Add headers to the Excel sheet
    #headers = ['Instance ID', 'Instance Type', 'Private IP', 'Public IP', 'CPU', 'Memory', 'Disk_Details', 'State', 'Launch Time']
    #ws.append(headers)

    # Add data to the Excel sheet
    for instance_info in ec2_inventory:
        ws.append([
            instance_info['Account'],
            instance_info['Region'],
            instance_info['Name'],

            instance_info['InstanceId'],
            instance_info['InstanceType'],
            instance_info['PrivateIpAddress'],
            instance_info['PublicIpAddress'],
            instance_info['cpu_details'],
            instance_info['memory_details'],
            instance_info['disk_details'],
            instance_info['State'],
            instance_info['LaunchTime'],
            instance_info['ProductTag'],
            instance_info['Tags'],
        ])

    # Save the Excel file
    wb.save(excel_file)

if __name__ == "__main__":
    
    wb = Workbook()
    ws = wb.active

    # Add headers to the Excel sheet
    headers = ['Account','Region','Name','Instance ID', 'Instance Type', 'Private IP', 'Public IP', 'CPU', 'Memory', 'Disk_Details', 'State', 'Launch Time','ProductTag','Tags']
    ws.append(headers)

    
    
    # Update with your role ARN and session name
    # To have the account as an array and loop through them to get all inventory
    
    
    #regions = ['eu-west-1']  # Update with your desired AWS region
    regions = ['eu-west-1','eu-west-2']
    #hcuk_accounts = ['217034714304']
    #hcuk_accounts = ['217034714304','489382767057','037256221545','757077737282','700448755547','282769123767','951740292612','453622666523','920321948046yes','146759720730yes','086609446728yes','586312944700','454697119165','557860412967','547343590405','389234147152','032960016182','099413669329']
    hcuk_accounts = ['920321948046','146759720730','086609446728']
    
    session_name = 'AssumedRoleSession'
    
    for hcuk_account in hcuk_accounts:
        for region in regions:

            role_arn     = "arn:aws:iam::" + hcuk_account + ":role/hcuk_terraform_deployment_role"
    
            ec2_inventory=[]
            ec2_inventory = assume_role_and_get_ec2_inventory(role_arn, session_name, region,hcuk_account)
            
            excel_file = 'ec2_inventory_Jan202501v1.xlsx'  # Update with your desired Excel file name
            export_to_excel(ec2_inventory, excel_file,wb,ws)

    #print(f"EC2 inventory exported to {excel_file}")
    # end
