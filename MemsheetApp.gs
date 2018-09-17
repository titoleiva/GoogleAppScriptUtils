MemsheetApp = {
  list: [],
  updateSheet: function(_name) {
  
  // get the sheet
   s = this.createSheet(_name);
   
   // check if the sheet exists
   for (i in this.list) {
      
      // if exists update it
      if (this.list[i].name == _name) {
         this.list[i] = s;
         return this.list[i];
      }
    }
    
    // if it doesn't exist add it
    this.list.push(s);
    return s;
  },
  getFirstSheet : function() {
   
    return this.list[0];
  },
  getSheet: function(_name) {
 
   // check if the sheet exists
   for (i in this.list) {
      
      // if exists get it
      if (this.list[i].name == _name) {         
         return this.list[i];
      }
    }
    
    // if it doesn't exist create it and add it
    s = this.createSheet(_name);
    this.list.push(s);
    return s;
  },
  createSheet: function(_name) {
      
    sheet = {
      sheet: SpreadsheetApp.getActive().getSheetByName(_name),
      name: _name,
      maxRow: SpreadsheetApp.getActive().getSheetByName(_name).getMaxRows(),
      maxCol: SpreadsheetApp.getActive().getSheetByName(_name).getMaxColumns(),
      rows: SpreadsheetApp.getActive().getSheetByName(_name).getRange(1,1,SpreadsheetApp.getActive().getSheetByName(_name).getMaxRows(), SpreadsheetApp.getActive().getSheetByName(_name).getMaxColumns()).getValues(),
      firstModifiedCell: [0,0],
      lastModifiedCell: [0,0],
      getActiveSheet: function() {
      
       return this.sheet;
      
      },
      getValues: function() {
      
       return this.rows;
      
      },
      getId: function() {
        return this.sheet.getId();
      },
      getLastRow: function() {
        return this.rows.length;
      },
      getLastColumn: function() {
        return this.rows[0].length;
      },
      getRow: function(r) {
      
        row = [];
        row.push(this.rows[r-1]);
        
        return row;
      },
      getColumn: function(col) {
      
        column = [];
        
        for (var i=0; i < this.rows.length; i++){
          column[i] = [];
          column[i].push(this.rows[i][col-1]);
       }
       return column;
      
        return this.rows[0].length;
      },
      getModifiedCells: function() {
      
        var cells = [];
        
        if (this.firstModifiedCell[0] == 0 || this.firstModifiedCell[1] == 0 || this.lastModifiedCell[0] == 0 || this.lastModifiedCell[1] == 0) {
         
          return 0;
        }
      
        // multiple cells
        for (var i = this.firstModifiedCell[0], r = 0; i <= this.lastModifiedCell[0]; r++, i++) {
           
           cells[r] = [];
           
           for (var j = this.firstModifiedCell[1], c = 0; j <= this.lastModifiedCell[1]; c++, j++) {
              
              var value = this.rows[i-1][j-1];
              cells[r][c] = value;
           }
        }
      
      return cells;
      
      },
      getCell: function(row, col) {
        
        if (!row) {
          row = col.substring(1);
          col = col.substring(0, 1);
        }
        if (isNaN(row)) {
          throw new Error("Multicell ranges not supported unless separating col and row in separate parameters");
        }
        
        var c = col;
        if (typeof col  === "string"){
          c = col.charCodeAt(0) - 65;
          // this supports 2 letters in col
          if (col.length > 1) {
            //"AB": 1 * (26) + 1 = 27 
            c = ( (c + 1) * ("Z".charCodeAt(0) - 64)) + (col.charCodeAt(1) - 65);
          }
        }
        
        c = parseInt(col)-1;
        if (this.maxCol < c) {
          this.maxCol = c;
        }
        
        var r = parseInt(row) - 1;
        if (this.maxRow < r) {
          this.maxRow = r;
        }
        
        if (!this.rows[r]) {
          this.rows[r] = [];
        }
        
        if (!this.rows[r][c]) {
          this.rows[r][c] = 0;
        }
        
        return {
          rows: this.rows,
          firstModifiedCell: this.firstModifiedCell,
          lastModifiedCell: this.lastModifiedCell,
          getValue: function() {
            return this.rows[r][c];
          },
          setValue: function(value) {
            this.rows[r][c] = value;
            
            if (this.firstModifiedCell[0] == 0 || this.firstModifiedCell[0] > row) {
               this.firstModifiedCell[0] = row;
            }
            
            if (this.firstModifiedCell[1] == 0 || this.firstModifiedCell[1] > col) {
               this.firstModifiedCell[1] = col;
            }
            
            if (this.lastModifiedCell[0] < row) {
               this.lastModifiedCell[0] = row;
            }
            
            if (this.lastModifiedCell[1] < col) {
               this.lastModifiedCell[1] = col;
            }
          }
        }
      }
    };
    return sheet;
  },
  flush: function() {
  
      for (i in this.list) {
      
        this.flushByIndex(i);
      } 
  },
  flushByName: function(_name) {
  
     for (i in this.list) {
      
      l = this.list[i];
      
      if (l.name == _name)
      
        this.flushByIndex(_name);
        return;
      } 
  },
  flushByIndex: function(index) {
  
    console.log(this.list);
    
    l = this.list[index];
      rowDiff = l.rows.length - Object.keys(l.rows).length;
      if (rowDiff > 0) {
        // insert empty rows at missing row entries
        emptyRow = [];
        for (c = 0; c < l.rows[0].length; c++) {
          emptyRow.push("");
        }
        for (j = 0; j < l.rows.length && rowDiff > 0; j++) {
          if (!l.rows[j]) {
            l.rows[j] = emptyRow;
            rowDiff--;
          }
        }
      }
      
      var cells = l.getModifiedCells();
    
    if (cells) {
      
      l.getActiveSheet().getRange(l.firstModifiedCell[0], l.firstModifiedCell[1], l.lastModifiedCell[0]-l.firstModifiedCell[0]+1, l.lastModifiedCell[1]-l.firstModifiedCell[1]+1).setValues(cells);
    }
      
      
    SpreadsheetApp.flush();
  }
}

