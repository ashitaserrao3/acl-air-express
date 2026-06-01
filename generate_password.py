import streamlit_authenticator as stauth
 
hashed_password = stauth.Hasher.hash("Welcome@123")
 
print(hashed_password)