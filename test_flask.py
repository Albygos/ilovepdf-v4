import app

client = app.app.test_client()

response = client.get('/hi/merge')
print("Request to /hi/merge:")
print(f"Status Code: {response.status_code}")
if response.status_code in (301, 302):
    print(f"Redirect Location: {response.headers.get('Location')}")
    
response2 = client.get('/hi/merge_pdf')
print("\nRequest to /hi/merge_pdf:")
print(f"Status Code: {response2.status_code}")
print(f"Body length: {len(response2.data)}")
