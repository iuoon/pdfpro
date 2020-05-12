from PIL import Image, ImageFilter
import numpy as np
from scipy import ndimage, misc

if __name__ == "__main__":
    rawfile = np.fromfile('E:\AQ\WD\zhongyuan.raw', "uint16")
    rawfile.shape = (1025,1025)
    misc.imsave("E:\AQ\WD\zhongyuan.png", rawfile)