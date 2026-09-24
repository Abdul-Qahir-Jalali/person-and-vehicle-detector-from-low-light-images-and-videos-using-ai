import cv2
import numpy as np

class ImageEnhancer:
    def __init__(self):
        # Create CLAHE object for contrast enhancement (low light)
        self.clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        
        # Sharpening kernel for deblurring
        self.sharpen_kernel = np.array([
            [-1, -1, -1],
            [-1,  9, -1],
            [-1, -1, -1]
        ])

    def enhance(self, image: np.ndarray) -> np.ndarray:
        """
        Enhance an image by fixing low light and blur, optimized for foggy conditions.
        """
        if image is None:
            return image

        # 1. Gamma Correction to reduce blinding headlight glare in night/fog conditions
        # A gamma > 1 darkens the midtones/highlights, pulling details out of the glare
        gamma = 1.5
        inv_gamma = 1.0 / gamma
        table = np.array([((i / 255.0) ** inv_gamma) * 255
                          for i in np.arange(0, 256)]).astype("uint8")
        glare_reduced = cv2.LUT(image, table)

        # 2. Noise reduction while keeping edges sharp (crucial for foggy images)
        smoothed = cv2.bilateralFilter(glare_reduced, 9, 75, 75)

        # 3. Low light & contrast enhancement using CLAHE
        lab = cv2.cvtColor(smoothed, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)
        l_enhanced = self.clahe.apply(l)
        lab_enhanced = cv2.merge((l_enhanced, a, b))
        enhanced_image = cv2.cvtColor(lab_enhanced, cv2.COLOR_LAB2BGR)
        
        # 3. Very mild sharpening (harsh sharpening creates blocky artifacts in fog)
        # Using a softer kernel instead of the harsh Laplacian
        soft_sharpen = np.array([
            [0, -0.5, 0],
            [-0.5, 3, -0.5],
            [0, -0.5, 0]
        ])
        sharpened_image = cv2.filter2D(enhanced_image, -1, soft_sharpen)
        
        return sharpened_image

enhancer = ImageEnhancer()
