import numpy as np
import trimesh
import os
import xlsxwriter

class STLGeometryToExcel:
    
    def __init__(self, file_path):
        self.file_path = file_path 
        self.mesh = trimesh.load(self.file_path)

    def calc_stl_bounding_box_volume(self):
        self.bounding_box = self.mesh.bounding_box.extents 
        bounding_box_volume = 1 
        for i in range(0,3): 
            bounding_box_volume = bounding_box_volume * self.bounding_box[i];
        return(bounding_box_volume)

    def calc_stl_geometry_data(self):
        self.stl_volume = self.mesh.volume

        self.bounding_box_volume = self.calc_stl_bounding_box_volume()

        self.height = self.bounding_box[2] # Height is the z-axis
        self.x_length = self.bounding_box[0]
        self.y_length = self.bounding_box[1]

        self.bottom_area = self.x_length * self.y_length

        self.surface_area = self.mesh.area

        self.stl_file_name = os.path.basename(self.file_path)

        return(self.stl_file_name)

    def single_stl_to_excel(self, excel_file):
        labels = ["Part Name", "Part Volume", "Bounding Box Volume", "Height", "X Length", "Y Length", "Surface Area", "Bottom Area"]
        stl_data = [self.stl_file_name, self.stl_volume, self.bounding_box_volume, self.height, self.x_length, self.y_length, self.surface_area, self.bottom_area]
        
        workbook = xlsxwriter.Workbook(excel_file)
        worksheet = workbook.add_worksheet()

        row = 0 
        column = 0
        for label in labels: 
            worksheet.write(row,column,label)
            column += 1 

        row = 1
        column = 0 
        for data in stl_data:
            worksheet.write(row,column,data)
            column +=1

        workbook.close()







file_path = "C:\CAD Files\Iso-Boundingbox Test\Screw algorithm test\scew_model_test.stl"
stlGeometryToExcel = STLGeometryToExcel(file_path)
stlGeometryToExcel.calc_stl_geometry_data()

stlGeometryToExcel.single_stl_to_excel("ReverseEngineeringTests.xlsx")



