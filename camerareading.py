import cv2
import numpy as np
from pytesseract import image_to_string
import pytesseract as pyt
import pandas as pd

# Initialize camera
cap = cv2.VideoCapture(0)  # 0 would make the program use the laptop's camera

# Set PATH to tesseract installation
pyt.pytesseract.tesseract_cmd = "C:\\Users\\sgeorge3\\AppData\\Local\\Tesseract-OCR\\tesseract.exe"

# Function to check if any valid words are found in the text from the image
def check_words(text, valid_words):
    found_words = [word for word in valid_words if word in text]
    return found_words

# Function to handle image rotation for better OCR accuracy
def rotate_image(image, angle):
    image_center = tuple(np.array(image.shape[1::-1]) / 2)
    rot_mat = cv2.getRotationMatrix2D(image_center, angle, 1.0)
    result = cv2.warpAffine(image, rot_mat, image.shape[1::-1], flags=cv2.INTER_LINEAR)
    return result

# Function to get movement data from the user, ensuring mandatory fields are not empty
def get_movement_data(movement):
    while True:
        Location = input("Please enter the location (could be the name of a person or company): ")
        if Location:
            break
        print("Location cannot be empty.")

    while True:
        shipping_Date = input("Enter the shipping date in MM/DD/YYYY format: ")
        if shipping_Date:
            break
        print("Shipping date cannot be empty.")

    while True:
        return_Date = input("Enter the return date in MM/DD/YYYY format: ")
        if return_Date:
            break
        print("Return date cannot be empty.")

    while True:
        shipping_Information = input("Enter the shipping information (e.g., FedEx tracking number): ")
        if shipping_Information:
            break
        print("Shipping information cannot be empty.")

    notes = input("Any additional comments: ")

    fedEx_number = "FedEx Tracking: " + shipping_Information
    return [Location, shipping_Date, return_Date, fedEx_number, notes]

# Function to update the Excel sheet based on the provided data and movement type
def update_excel(data_list, movement):
    file_path = 'testsheet.xlsx'
    sheet_name = "Sheet1"

    # Reading the Excel file into a dataframe
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    first_item = data_list[0]
    match_found = False

    for index in df.index:
        if df.iloc[index, 0] == first_item:
            if movement == "2":  # If returning, clear the cells next to the ID
                df.iloc[index, 1:6] = [None] * 5
                print(f"System {first_item} has been returned and fields have been cleared.")
            df.iloc[index, :len(data_list)] = data_list
            match_found = True
            break

    if not match_found:
        new_row = pd.DataFrame([[data_list[0]] + [None]*(len(df.columns)-1)], columns=df.columns)
        df = pd.concat([df, new_row], ignore_index=True)
        new_index = df.index[-1]
        df.iloc[new_index, :len(data_list)] = data_list

    df.to_excel(file_path, sheet_name=sheet_name, index=False)
    print(f"Excel file has been updated with the data: {data_list}")

    # Release the camera and close the OpenCV windows, then exit the program
    cap.release()
    cv2.destroyAllWindows()
    exit()

# Read the Excel file and extract the first column into a list
file_path = 'testsheet.xlsx'
sheet_name = "Sheet1"
df = pd.read_excel(file_path, sheet_name=sheet_name)
list_ID = df.iloc[:, 0].tolist()

# Users we will be detecting for
valid_words = ['OEMDSK57', 'Welcome', '#WeareLenovo', 'batman', 'spiderman']

while True:
    # Capturing the frames
    ret, frame = cap.read()

    # Check if frame is captured successfully
    if not ret:
        print("Error: failed to capture the frame")
        break

    # Display captured frame
    cv2.imshow('Facial Recognition', frame)

    # Keypress
    key = cv2.waitKey(1)
    if key == ord('q'):  # Quit
        break
    elif key == ord('c'):  # Capture image for facial recognition simulation
        for angle in [0, 90, 180, 270]:  # Check different orientations
            rotated_frame = rotate_image(frame, angle)
            text = image_to_string(rotated_frame)
            found_words = check_words(text, valid_words)

            if found_words:
                print(f"Authorized user detected: {', '.join(found_words)}")
                cv2.destroyAllWindows()  # Close the OpenCV window
                barcode = input("Please scan the barcode...")

                # Modify the barcode input to uppercase the first five characters
                barcode = barcode[:5].upper() + barcode[5:]

                while True:
                    if barcode in list_ID:
                        print("Valid system detected: " + barcode)

                        movement = input("Please decide which of the following you want \n 1.Shipping \n 2.Returning \n 3.Taking \n")
                        if movement in ["1", "2", "3"]:
                            data = get_movement_data(movement)
                            data_list = [barcode] + data
                            update_excel(data_list, movement)
                            break
                        else:
                            print("Invalid movement option.")
                    else:
                        print("This system does not exist, please try again or check if this system is valid.")
                        break                     
            break
        else:
            print("No authorized user detected, please try again if you think you may have clearance.")

cap.release()
cv2.destroyAllWindows()
