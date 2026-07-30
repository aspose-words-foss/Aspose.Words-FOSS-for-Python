"""Image-related enums."""


class ImageType:
    """Specifies the type (format) of an image in a document."""

    NO_IMAGE = 0
    UNKNOWN = 1
    EMF = 2
    WMF = 3
    PICT = 4
    JPEG = 5
    PNG = 6
    BMP = 7
    EPS = 8
    WEB_P = 9
    GIF = 10
    SVG = 11
    TIFF = 12


_MIME_TO_IMAGE_TYPE: dict[str, int] = {
    "image/png": ImageType.PNG,
    "image/jpeg": ImageType.JPEG,
    "image/jpg": ImageType.JPEG,
    "image/gif": ImageType.GIF,
    "image/bmp": ImageType.BMP,
    "image/tiff": ImageType.TIFF,
    "image/x-emf": ImageType.EMF,
    "image/x-wmf": ImageType.WMF,
    "image/svg+xml": ImageType.SVG,
    "image/webp": ImageType.WEB_P,
    "image/eps": ImageType.EPS,
}

_IMAGE_TYPE_TO_MIME: dict[int, str] = {
    ImageType.PNG: "image/png",
    ImageType.JPEG: "image/jpeg",
    ImageType.GIF: "image/gif",
    ImageType.BMP: "image/bmp",
    ImageType.TIFF: "image/tiff",
    ImageType.EMF: "image/x-emf",
    ImageType.WMF: "image/x-wmf",
    ImageType.SVG: "image/svg+xml",
    ImageType.WEB_P: "image/webp",
    ImageType.EPS: "image/eps",
}
