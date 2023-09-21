import os
import requests
import pandas as pd


def organize_reports(dataframe, obj, obj_percent):
    """
    Sorts the reports in the onedrive directory into folders for each advisor and client.

    Args:
        dataframe: A Pandas DataFrame with the relation "cliente", "conta", "assessor".
    """
    # Get the user's home directory
    home_dir = os.path.expanduser("~")

    # Construct the path to the Downloads folder and the "planilhas_btg" folder
    download_folder = os.path.join(home_dir, "Downloads")
    btg_folder = os.path.join(download_folder, "planilhas_btg")

    # Get the list of folders in the btg_folder
    folder_list = os.listdir(btg_folder)

    # Get authentication token
    auth_response = requests.post(
        "https://login.microsoftonline.com/4d2a63d3-c3a5-415e-9e8c-3659deb77b56/oauth2/v2.0/token",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={
            "client_id": "bbd8003e-f351-486e-aa03-c3e09aeed98b",
            "client_secret": "Ztf8Q~qU4EyNjxOPBjOqKKJgsROLDFeW34wjxcUc",
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=10,
    )

    print(auth_response.json())
    # Get access token
    access_token = auth_response.json()["access_token"]

    # Get user's sites
    sites_response = requests.get(
        "https://graph.microsoft.com/v1.0/users/",
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=10,
    )

    # Get user's sites data
    sites_data = sites_response.json()

    # Make a list of the files in the downloaded directory
    files = os.listdir(btg_folder)

    # Create a set of unique advisors
    advisors = set(dataframe["assessor"])

    # Iterate over each advisor
    bar = 0
    for advisor in advisors:
        advisor_normalized = advisor.lower().replace(" ", "_")
        advisor_dir = os.path.join(btg_folder, advisor_normalized)

        # Filter the dataframe for rows where the advisor matches
        advisor_data = dataframe[dataframe["assessor"] == advisor]

        # Iterate over each client
        for _, row in advisor_data.iterrows():
            client_dir = os.path.join(advisor_dir, str(row["conta"]))
            # List all files in the download directory
            files = os.listdir(btg_folder)

            # Iterate over the files and move them to the client's folder if the client ID is in the filename
            for filename in files:
                bar += 0.17
                obj.UpdateBar(bar + (len(files) * 0.5), len(files))
                obj_percent.update(
                    str(int(((bar + (len(files) * 0.5)) / len(files)) * 100)) + "%"
                )
                if str(row["conta"]) in filename:
                    # source_path = os.path.join(btg_folder, filename)
                    destination_path = os.path.join(client_dir, filename)
                    # Replace backslashes with forward slashes
                    destination_path = destination_path.replace("\\", "/")

                    # String to remove
                    string_to_remove = "C:/Users/Workspace/Downloads"

                    # Replace the string with an empty string
                    destination_path_fixed = destination_path.replace(
                        string_to_remove, ""
                    )

                    print(destination_path)
                    # Read the file as binary data
                    with open(os.path.join(btg_folder, filename), "rb") as f:
                        file_data = f.read()

                    # Send the file to OneDrive
                    upload_url = f"https://graph.microsoft.com/v1.0/users/4d4bf836-7e9e-455f-be46-5cb9c5cc729f/drive/root:{destination_path_fixed}:/content"
                    headers = {
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/pdf",  # Set content type to JSON
                    }

                    # Send the file to OneDrive with the updated headers and JSON payload
                    response = requests.put(
                        upload_url, headers=headers, data=file_data, timeout=10
                    )
            obj.UpdateBar(1, 1)
            obj_percent.update("Concluído")
