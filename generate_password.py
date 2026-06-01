import streamlit_authenticator as stauth
 
passwords = [
    "pradeep@123",
    "krishna@123",
    "shivpratap@123",
    "manoj@123",
    "ram@123",
    "sujeet@123",
    "vinod@123"

]
 
for password in passwords:
    print(stauth.Hasher.hash(password))