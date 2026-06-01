import streamlit_authenticator as stauth
 
passwords = [
    "ashita@123",
    "arvin@123",
    "apu@123",
    "srinivas@123",
    "devendra@123",
    "pramod@123",
    "chandrasekar@123",
    "sanjay@123",
    "dinesh@123",
    "shailendra@123"
]
 
for password in passwords:
    print(stauth.Hasher.hash(password))