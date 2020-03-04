from controller import Controler_relance_client
from gui_relance_facture import Gui_relance_client

view = Gui_relance_client()
app = Controler_relance_client(view)
view.show()
