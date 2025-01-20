from pylibdmtx.pylibdmtx import encode
from PIL import Image
encoded = encode('hello world'.encode('utf8'))
img = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
img.save('dmtx.png')