import requests
import base64
import sys

def main():
    if len(sys.argv) < 2:
        print("Usage: python test_client.py <image_path>")
        sys.exit(1)

    image_path = sys.argv[1]
    url = "http://127.0.0.1:8000/api/v1/detect"

    with open(image_path, "rb") as f:
        files = {"file": (image_path, f, "image/jpeg")}
        print(f"Sending request to {url}...")
        response = requests.post(url, files=files)

    if response.status_code == 200:
        data = response.json()
        print("Success:", data["success"])
        print("Message:", data["message"])
        print("Person Count:", data["person_count"])
        print("Vehicle Count:", data["vehicle_count"])
        
        # Save the annotated image
        if data.get("annotated_image_base64"):
            with open("annotated_output.jpg", "wb") as f_out:
                f_out.write(base64.b64decode(data["annotated_image_base64"]))
            print("Saved annotated image to annotated_output.jpg")
    else:
        print(f"Error: {response.status_code}")
        print(response.text)

if __name__ == "__main__":
    main()
