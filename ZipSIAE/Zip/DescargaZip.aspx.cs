using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using BusinessEntity;
using BusinessLogic;

namespace Zip
{
    public partial class DescargaZip : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnDescargar_Click(object sender, EventArgs e)
        {

           
            String IdNodo = txtNodo.Text;
            EntidadDetalleBE rutaVirtualTemporalBE = new EntidadDetalleBE();
            rutaVirtualTemporalBE.Entidad.IdEntidad = "CONF";
            rutaVirtualTemporalBE.IdValor = "RUTA_VIRT_TEMP";
            rutaVirtualTemporalBE = EntidadDetalleBL.ListarEntidadDetalle(rutaVirtualTemporalBE)[0];

            ZipBL zip = new ZipBL();

            zip.DescargarZip(IdNodo);

            String nombreCarpetaZip = IdNodo + ".zip";

            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "impresion", "window.open('" + rutaVirtualTemporalBE.ValorCadena1 + "/" + nombreCarpetaZip + "');", true);


        }
    }
}