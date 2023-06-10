import os

import pypff
import unicodecsv as csv

def main(pst_file):
    """
    The main function opens a PST and calls functions to parse and report data from the PST
    :param pst_file: A string representing the path to the PST file to analyze
    :param report_name: Name of the report title (if supplied by the user)
    :return: None
    """
    print("Opening PST for processing...")
    pst_file_f = open(pst_file, "rb")
    opst = pypff.open_file_object(pst_file_f)
    root = opst.get_root_folder()

    print("Starting traverse of PST structure...")
    folderTraverse(root)


def makePath(file_name):
    """
    The makePath function provides an absolute path between the output_directory and a file
    :param file_name: A string representing a file name
    :return: A string representing the path to a specified file
    """
    folder = os.path.join(os.getcwd(), 'output')
    os.makedirs(folder, exist_ok=True)
    return os.path.abspath(os.path.join(folder, file_name))


def folderTraverse(base):
    """
    The folderTraverse function walks through the base of the folder and scans for sub-folders and messages
    :param base: Base folder to scan for new items within the folder.
    :return: None
    """
    for folder in base.sub_folders:
        if folder.number_of_sub_folders:
            folderTraverse(folder) # Call new folder to traverse:
        checkForMessages(folder)


def checkForMessages(folder):
    """
    The checkForMessages function reads folder messages if present and passes them to the report function
    :param folder: pypff.Folder object
    :return: None
    """
    exlude_list = {'Deleted Items','Detected Items','Calendar', 'Junk Email', 'Recipient Cache', 'חגים בישראל'}
    if folder.name in exlude_list:
        return
    print("Processing Folder: " + folder.name)
    message_list = []
    for message in folder.sub_messages:
        message_dict = processMessage(message)
        message_list.append(message_dict)
    folderReport(message_list, folder.name)


def processMessage(message):
    """
    The processMessage function processes multi-field messages to simplify collection of information
    :param message: pypff.Message object
    :return: A dictionary with message fields (values) and their data (keys)
    """
    return {
        "subject": message.subject,
        "sender": message.sender_name,
        "header": message.transport_headers,
        "body": message.plain_text_body,
        "creation_time": message.creation_time,
        "submit_time": message.client_submit_time,
        "delivery_time": message.delivery_time,
        "attachment_count": message.number_of_attachments,
    }


def folderReport(message_list, folder_name):
    """
    The folderReport function generates a report per PST folder
    :param message_list: A list of messages discovered during scans
    :folder_name: The name of an Outlook folder within a PST
    :return: None
    """
    if not len(message_list):
        print("Empty message not processed")
        return

    messages_with_bodies = []
    for m in message_list:
        if m['body']:
            data = m['body']
            if isinstance(data, str):
                data = data.encode('utf-8-sig', 'ignore')
            m['body'] = data.decode()
            messages_with_bodies.append(m)
    print(f'Total messages with body in {folder_name}:{len(messages_with_bodies)}')
    if messages_with_bodies:
        # CSV Report
        fout_path = makePath("folder_report_" + folder_name + ".csv")
        fout = open(fout_path, 'wb')
        # fout.write(bytes(u'\ufeff', encoding='utf-8'))
        header = ['creation_time', 'submit_time', 'delivery_time', 'sender', 'subject', 'body', 'attachment_count']
        csv_fout = csv.DictWriter(fout, fieldnames=header, extrasaction='ignore', encoding='utf-8-sig', errors='ignore')
        csv_fout.writeheader()
        csv_fout.writerows(messages_with_bodies)
        fout.close()

        # HTML Report Prep
        body_out = open(makePath(f"message_body_{folder_name}.txt"), 'wb')
        # dict_keys(['subject', 'sender', 'header', 'body', 'creation_time', 'submit_time', 'delivery_time', 'attachment_count'])
        for m in messages_with_bodies:
            if m['body']:
                data = m['body']
                if isinstance(data, str):
                    data = data.encode('utf-8', 'ignore')
                body_out.write(data + b"\n\n")
        body_out.close()


if __name__ == "__main__":
    file = r'test.ost'
    print('Starting Script...')
    main(file)
    print('Script Complete')