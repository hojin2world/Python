import hashlib

# 사용자가 입력한 비밀번호
password = "590215"

# SHA-512 해싱
hashed_password = hashlib.sha256(password.encode()).hexdigest()

print(hashed_password)