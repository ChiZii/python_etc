# We
# 한글 형태소 분석 모듈
# kiwipiepy
# https://github.com/bab2min/kiwipiepy
from kiwipiepy import Kiwi

kiwi = Kiwi()
print(kiwi.tokenize('형태소를 구분을 지어서 보여주는 모듈'))

print(kiwi.space('띄어쓰기를자동으로해주는기능'))
