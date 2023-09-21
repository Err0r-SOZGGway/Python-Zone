import random
omikuji = random.randint(1,20) # 1から20の数値をランダムに生成
#print(omikuji)

if omikuji == 1:
    print("大吉")
elif omikuji <= 4: # 2,3,4
    print("中吉")
elif omikuji <= 8: # 5,6,7,8
    print("小吉")
elif omikuji <= 12: # 9,10,11,12
    print("吉")
elif omikuji <= 15: # 13,14,15
    print("末吉")
elif omikuji <= 18: # 16,17,18
    print("凶")
else:              # 19,20
    print("大凶")