import qrcode

def make_qr(name, date):
    qr = qrcode.QRCode(
    version = 1,
    error_correction = qrcode.constants.ERROR_CORRECT_L,
    box_size = 10,
    border = 4)

    qr.add_data(f"{name} has earned this certificate in {date}")
    qr.make(fit=True)

    qr_image = qr.make_image(fill_color = "black", back_color = "white")

    return qr_image