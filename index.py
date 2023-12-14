
from datetime import datetime
import win32com.client
import matplotlib.pyplot as plt
import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import wx
import glob
import re
from PyPDF2 import PdfReader, PdfWriter, Transformation, PageObject, PaperSize, PdfMerger
from PyPDF2.generic import RectangleObject
user_files = list()


class DXF2IMG(object):
    default_img_format = '.pdf'
    default_img_res = 300
    default_bg_color = '#FFFFFF'
    files_per_batch = 20

    def __init__(self, gui):
        self.gui = gui

    def convert_dxf2img(self, names, img_format=default_img_format, img_res=default_img_res, clr=default_bg_color):
        pdf_merger = PdfMerger()
        total_files = len(names)
        batch_count = 0
        success = True
        try:
            for index, name in enumerate(names, start=1):
                doc = ezdxf.readfile(name)
                msp = doc.modelspace()
                auditor = doc.audit()
                if len(auditor.errors) != 0:
                    raise Exception(f"DXF DAÑADO --> {name}")
                else:
                    fig = plt.figure()
                    ax = fig.add_axes([0, 0, 1, 1])
                    ctx = RenderContext(doc)
                    ctx.set_current_layout(msp)
                    ezdxf.addons.drawing.properties.MODEL_SPACE_BG_COLOR = clr
                    out = MatplotlibBackend(ax)
                    Frontend(ctx, out).draw_layout(msp, finalize=True)
                    img_name = re.findall("(\S+)\.", name)
                    pdf_file = ''.join(img_name) + img_format
                    fig.savefig(pdf_file, dpi=img_res)
                    # Apply PDF page formatting
                    self.format_pdf_pages(pdf_file)
                    pdf_merger.append(pdf_file)
                    if index % self.files_per_batch == 0 or index == total_files:
                        batch_count += 1

                        pdf_merger = PdfMerger()
                        progress_percentage = int(index / total_files * 100)
                        self.gui.update_progress_bar(progress_percentage)
        except Exception as e:
            success = False
         
            wx.MessageBox('Error durante la conversión: {e}"', 'Error',
                      wx.OK | wx.ICON_ERROR)
        if success:
            wx.CallAfter(self.gui.show_success_dialog)

    def format_pdf_pages(self, pdf_file):
        A4_w = PaperSize.A4.width
        A4_h = PaperSize.A4.height

        pdf_reader = PdfReader(pdf_file)
        pdf_writer = PdfWriter()

        for page in pdf_reader.pages:
            h = float(page.mediabox.height)
            w = float(page.mediabox.width)
            scale_factor = min(A4_h / h, A4_w / w)
            transform = Transformation().scale(
                scale_factor, scale_factor).translate(0, A4_h / 3)
            page.add_transformation(transform)
            page.cropbox = RectangleObject((0, 0, A4_w, A4_h))
            page_A4 = PageObject.create_blank_page(width=A4_w, height=A4_h)
            page.mediabox = page_A4.mediabox
            page_A4.merge_page(page)
            pdf_writer.add_page(page_A4)

        with open(pdf_file, "wb") as output_pdf:
            pdf_writer.write(output_pdf)



class Interfaz(wx.Frame):
    def __init__(self):
        super().__init__(None, title='DXF Converter', size=(400, 400),
                         style=wx.MINIMIZE_BOX | wx.RESIZE_BORDER | wx.SYSTEM_MENU |
                         wx.CAPTION | wx.CLOSE_BOX | wx.CLIP_CHILDREN)
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)
        # Agregar una barra de progreso
        self.progress_bar = wx.Gauge(
            panel, range=100, size=(300, 25), style=wx.GA_HORIZONTAL)
        sizer.Add(self.progress_bar, 0, wx.CENTER | wx.ALL, 10)

        # Agregar un cuadro de texto para mostrar el total de archivos seleccionados
        self.total_files_text = wx.StaticText(panel, label='Total de archivos seleccionados: 0')
        sizer.Add(self.total_files_text, 0, wx.CENTER | wx.ALL, 10)

        folder_btn = wx.Button(
            panel, label='Seleccione la Carpeta', size=(150, 30))
        folder_btn.Bind(wx.EVT_BUTTON, self.on_open_folder)
        sizer.Add(folder_btn, 0, wx.CENTER | wx.ALL, 10)

        convert_btn = wx.Button(
            panel, label='Convertir archivos', size=(150, 30))
        convert_btn.Bind(wx.EVT_BUTTON, self.on_convert)
        convert_btn.SetBackgroundColour(wx.Colour(0, 150, 0))
        sizer.Add(convert_btn, 0, wx.CENTER | wx.ALL, 10)

        close_btn = wx.Button(panel, label='Cerrar', size=(150, 30))
        close_btn.Bind(wx.EVT_BUTTON, self.on_close)
        close_btn.SetBackgroundColour(wx.Colour(250, 0, 0))
        sizer.Add(close_btn, 0, wx.CENTER | wx.ALL, 10)

        panel.SetSizer(sizer)
        self.Show()

    def update_total_files_text(self):
        total_files = len(user_files)
        self.total_files_text.SetLabel(f'Total de archivos seleccionados: {total_files}')

    def update_dxf_listing(self, folder_path):
        self.current_folder_path = folder_path
        dxfs = glob.glob(folder_path + '/*.dxf')
        if not dxfs:
            wx.MessageBox('No se han encontrado DXF en esta carpeta',
                          'Not Found', wx.OK | wx.ICON_ERROR)
        index = 0
        user_files.clear()  # Limpiar la lista antes de agregar nuevos archivos
        for dxf in dxfs:
            user_files.append(dxf)
        self.update_total_files_text()  # Actualizar el cuadro de texto


    def update_progress_bar(self, value):
        self.progress_bar.SetValue(value)

    def on_open_folder(self, event):
        title = "Selecciona un directorio:"
        dlg = wx.DirDialog(self, title, style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            user_folder2 = dlg.GetPath()
            self.update_dxf_listing(user_folder2)
        dlg.Destroy()

    def on_convert(self, event):
        global user_files
        first = DXF2IMG(self)
        if user_files:
            first.convert_dxf2img(user_files)
        else:
            wx.MessageBox('Selecciona una carpeta',
                          'No seleccionado', wx.OK | wx.ICON_ERROR)

    def on_close(self, event):
        self.Close()

    def show_success_dialog(self):
        wx.MessageBox('Conversión exitosa', 'Éxito',
                      wx.OK | wx.ICON_INFORMATION)


if __name__ == '__main__':
    app = wx.App()
    frame = Interfaz()
    app.MainLoop()