# main file

from app.fileRegistrationDb import Fileregister
from app.margePdfFiles import MargePdfFiles


manager = MargePdfFiles()

print(manager.Duplicatefilesfinder())