import cognitive_face as CF
from global_variables import personGroupId
import sys

Key = '7af26ce4f0db46a58f2280df501f27c3'
CF.Key.set(Key)
BASE_URL = 'https://westcentralus.api.cognitive.microsoft.com/face/v1.0'  # Replace with your regional Base URL
CF.BaseUrl.set(BASE_URL)

personGroups = CF.person_group.lists()
for personGroup in personGroups:
    if personGroupId == personGroup['personGroupId']:
        print (personGroupId + " already exists.")
        sys.exit()

res = CF.person_group.create(personGroupId)
print(res)
