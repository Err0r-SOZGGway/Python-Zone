import random

Hero_hp = 400
Satan_hp = 500

while Satan_hp > 0:
  hit = random.randint(10,50)
  print("魔王に" + str(hit) + "のダメージを与えた！")
  Satan_hp -= hit
  hit = random.randint(50,200)
  print("勇者に" + str(hit) + "のダメージ！")
  Hero_hp -= hit
  print("勇者のHP:" + str(Hero_hp))
  print("魔王のHP:" + str(Satan_hp))
  if Hero_hp < 0:
    break
  elif Satan_hp < 0:
    pass
  else:
    heal = random.randint(0,150)
    print("勇者のHPが" + str(heal) + "回復した！")
    Hero_hp += heal
    print("勇者のHP:" + str(Hero_hp))

if Satan_hp < 0 and Hero_hp < 0:
  print("勇者と魔王は相打ちになった・・・")
elif Hero_hp < 0 and Satan_hp > 0:
  print("勇者はしんでしまった！")
else:
  print("魔王を倒した！")