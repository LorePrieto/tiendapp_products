module TiendappProducts
  class Import
    require 'roo'

    def self.create_products(path)
      xlsx = Roo::Spreadsheet.open(path)
      if !self.is_not_a_correct_excel(xlsx)
        return "Buyyaaah!!"
      end
    end

    def self.is_not_a_correct_excel(xlsx)
      #Has the correct tabs and number of them
      if xlsx.sheets.count != 6
        return "Deben ser 6 hojas en el excel"
      end
      if !(xlsx.sheets.include? "Productos")
        return "No hay una hoja de Productos"
      end
      if !(xlsx.sheets.include? "Propiedades")
        return "No hay una hoja de Propiedades"
      end
      if !(xlsx.sheets.include? "Opciones")
        return "No hay una hoja de Opciones"
      end
      if !(xlsx.sheets.include? "Variantes")
        return "No hay una hoja de Variantes"
      end
      if !(xlsx.sheets.include? "Ubicaciones")
        return "No hay una hoja de Ubicaciones"
      end
      if !(xlsx.sheets.include? "Stock")
        return "No hay una hoja de Stock"
      end
      #Has the correct headers in the tabs
      if xlsx.sheet("Productos").row(1) != ["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías" ]
        return "Los headers en la hoja Productos no son correctos"
      end
      if xlsx.sheet("Propiedades").row(1) != ["ID Producto", "Propiedad", "Valor"]
        return "Los headers en la hoja Propiedades no son correctos"
      end
      if xlsx.sheet("Opciones").row(1) != ["ID Producto", "Opción", "Valores" ]
        return "Los headers en la hoja Opciones no son correctos"
      end
      if xlsx.sheet("Variantes").row(1) != ["ID", "ID Producto", "Opciones", "SKU", "Precio", "Peso", "Altura", "Longitud", "Profundidad"]
        return "Los headers en la hoja Variantes no son correctos"
      end
      if xlsx.sheet("Ubicaciones").row(1) != ["ID", "Nombre", "Nombre Interno", "Calle", "Ciudad", "Calle de referencia", "Código Postal", "Teléfono", "País", "Región", "Activa", "Por defecto", "Backorderable", "Propagar por todas las variantes"]
        return "Los headers en la hoja Ubicaciones no son correctos"
      end
      if xlsx.sheet("Stock").row(1) != ["ID Producto", "ID Variante", "Cantidad", "ID Ubicación", "Backorderable"]
        return "Los headers en la hoja Stock no son correctos"
      end
      return false
    end
  end
end
