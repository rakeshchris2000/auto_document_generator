import requests

SLACK_DEFAULT_ENDPOINT = 'http://localhost:14300/actors/uchat/send'

def send_notification(slack_channels, document_link, slack_endpoint=SLACK_DEFAULT_ENDPOINT):
    # Compose the message with the document link inserted properly
    message = (
        f"Hi Team,\n\n"
        "Please find the Delivery+ Deployment Safety MBR document for the current month below:\n\n"
        f"<{document_link}|Click here to view the document>\n\n"
        "Thanks!"
    )

    status_msgs = []
    for channel in slack_channels:
        request_body = {
            "channel": channel,
            "text": message
        }
        try:
            response = requests.post(url=slack_endpoint, json=request_body)
            response.raise_for_status()
            status_msg = response.json()
        except requests.exceptions.RequestException as e:
            status_msg = {"error": str(e)}
        except ValueError:
            status_msg = {"error": "Failed to parse response"}

        status_msgs.append({channel: status_msg})

    return status_msgs
