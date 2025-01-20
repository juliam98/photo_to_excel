from PIL import Image

def _rgb_str_format(rgb_val:list[int])->str:
    """
    Adjust the RGB values to the format used in Excel
    """
    rgb_val = [int(val) for val in rgb_val]  # Convert elements to integers
    rgb_formatted = f"FF{rgb_val[0]:02X}{rgb_val[1]:02X}{rgb_val[2]:02X}"
    return rgb_formatted
    
def _remove_alpha(rgb_tuple):
    """
    Remove the alpha channel from the RGB tuple
    """
    return rgb_tuple[:3] 

def _image_size_ratio(img:Image):
    """
    Calculate the ratio of the image size
    """
    ratio = img.size[0] / img.size[1]
    return ratio