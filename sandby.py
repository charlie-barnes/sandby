#!/usr/bin/env python
#-*- coding: utf-8 -*-

#Sandby Moth Recorder - A species recording tool
#This application is free software; you can redistribute
#it and/or modify it under the terms of the GNU General Public License
#defined in the COPYING file

#Copyright (C) 2009 Charlie Barnes. All rights reserved.

import os
import gtk
from sqlite3 import dbapi2 as sqlite
import gobject
import time
import subprocess
import csv
import cairo
import pango
import calendar
from datetime import date
import sys
from operator import itemgetter
import re
import logging
import string
import shutil

     
LOG_FILENAME = 'sandby.log'
logging.basicConfig(filename=LOG_FILENAME,level=logging.DEBUG)

try:
    Raise()
    import goocanvas
except:
    logging.debug("goocanvas module not loaded - disabling graphing")
        
DATE_SEPARATOR = "/"

checklists_data = []

class PrintData:
    text = None
    layout = None
    page_breaks = None

page_setup = None
settings = None
active_prints = []
stats_window_created = False

class sandbyInitialize:

    def __init__(self):
        #The version of Sandby Recorder based on last modified time of sys.argv[0]
        self.version = "Release 0." + time.strftime("%Y%m%d", time.gmtime(os.path.getmtime(sys.argv[0])))

        #Set up the program directories
        checklist_files = ["_bbm_template.db"]

        #Load the Glade widget tree
        self.glade = gtk.Builder()
        self.glade.add_from_file('sandby.glade')

        #The loaded database
        self.database_filename = None

        #Database handles
        self.con = None
        self.cur = None

        #Setup the main window
        self.main_window = self.glade.get_object("main_window")
        self.main_window.connect("destroy", self.quit)
        self.main_window.connect("window-state-event", self.on_window_state_change)
        self.window_in_fullscreen = False

        #The column headings
        self.column_headings = ['Code', 'Common Name', 'Scientific Name', 'Authority', 'Family', 'Count']

        #The database treeview
        self.database_table_rows = gtk.ListStore(str, str, str, str, str, int)
        self.database_table = self.glade.get_object("treeview2")
        self.database_table.set_headers_visible(False)
        self.database_table.set_reorderable(True)
        self.database_table_rows.set_sort_column_id(0, gtk.SORT_ASCENDING)
        
        selection = self.database_table.get_selection()
        selection.set_mode(gtk.SELECTION_MULTIPLE)

        #Set up the columns
        cell = gtk.CellRendererText()
        index=0

        for fieldDesc in self.column_headings:
            if fieldDesc == "Scientific Name":
                col = gtk.TreeViewColumn(fieldDesc, cell, markup=index)
            elif fieldDesc == "Code":
                col = gtk.TreeViewColumn(fieldDesc, cell, text=index)
                self.database_table_rows.set_sort_func(index, self.sort_bf)
            else:
                col = gtk.TreeViewColumn(fieldDesc, cell, text=index)

            self.database_table.append_column(col)
            self.database_table.get_column(index).set_resizable(True)
            col.set_sort_column_id(index)
            index+=1
                       
        self.database_table.set_model(self.database_table_rows)
 
        self.glade.get_object("toolbutton4").set_expand(True)

        completion = gtk.EntryCompletion()
        completion.connect("match-selected", self.on_completion_match)
        completion.set_text_column(0)
        completion.set_minimum_key_length(3)
        completion.set_match_func(self.completion_match, 0)     
        completion.set_model(gtk.ListStore(str))
        self.glade.get_object("entry6").set_completion(completion)
        
        throbber = self.glade.get_object("image2")
        animation = gtk.gdk.PixbufAnimation("loader.gif")
        throbber.set_from_animation(animation)
        self.glade.get_object("eventbox2").modify_bg(gtk.STATE_NORMAL, gtk.gdk.color_parse("#ffffff"))

        #Loop through our checklist_files 
        for file in checklist_files:
                try:
                    connection = sqlite.connect(file)
                    cursor = connection.cursor()
                    cursor.execute("Select checklist FROM meta")
                    checklist = cursor.fetchone()
                    checklists_data.append([checklist[0], file])
                except:
                    logging.debug(file + " doesn't appear to be a valid database file")

class sandbyActions(sandbyInitialize):
    def __init__(self):
        sandbyInitialize.__init__(self)
        self.connect_object()
        self.main_window.show()

        #Load (or create if it doesn't exist) the specified file if we have one
        if len(sys.argv) > 1:
            if os.path.exists(sys.argv[1]):
                self.database_filename = os.path.abspath(sys.argv[1])
                self.load_tables()
            else:
                self.create_database(sys.argv[1])

    def connect_object(self):
        mdict={"new_database":self.new_database,
               "open_database":self.open_database,
               "close_database":self.close_database,
               "save_database_as":self.save_database_as,
               "about":self.about,
               "import_csv":self.import_csv,
               "export_csv":self.export_csv,
               "switch_fullscreen":self.switch_fullscreen,
               "switch_toolbar":self.switch_toolbar,
               "add_data":self.add_data,
               "edit_data":self.edit_data,
               "do_data":self.do_data,
               "year_change":self.year_change,
               "month_change":self.month_change,
               "dataset_properties":self.dataset_properties,
               "dataset_properties_close":self.dataset_properties_close,
               "dataset_properties_hide":self.dataset_properties_hide,
               "day_change":self.day_change,
               "day_change_double_click":self.day_change_double_click,
               "toggle_calendar":self.toggle_calendar,
               "window_resize":self.window_resize,
               "do_print":self.do_print,
               "do_page_setup":self.do_page_setup,                                                          
               "show_statistics":self.show_statistics,
               "stats_close":self.stats_close,
               "stats_delete":self.stats_delete,
               "statistics_window_popup_menu":self.statistics_window_popup_menu,
               "main_window_popup_menu":self.main_window_popup_menu,
               "statistics_copy":self.statistics_copy,
               "statistics_select_all":self.statistics_select_all,
               "statistics_select_none":self.statistics_select_none,
               "main_window_copy":self.main_window_copy,
               "main_window_select_all":self.main_window_select_all,
               "main_window_select_none":self.main_window_select_none,
               "main_window_delete":self.main_window_delete,
               "main_window_edit":self.main_window_edit,
               "update_autocomplete":self.update_autocomplete,
               "plot_graph":self.plot_graph,
               "focus_in_species":self.focus_in_species,
               "focus_out_species":self.focus_out_species,
               "focus_in_count":self.focus_in_count,
               "focus_out_count":self.focus_out_count,
               "changed_count":self.changed_count,
               "del_key":self.del_key,
               "quit":self.quit
              }
        self.glade.connect_signals(mdict)

    def focus_in_species(self, widget, var):
        if widget.get_text() == "Species":
            widget.set_text("")

    def focus_out_species(self, widget, var):
        if widget.get_text() == "":
            widget.set_text("Species")
            
    def focus_in_count(self, widget, var):
        if widget.get_text() == "Count":
            widget.set_text("")

    def focus_out_count(self, widget, var):
        if widget.get_text() == "":
            widget.set_text("Count")

    def changed_count(self, widget, event):
        if not re.match ("([0-9.,z]|BackSpace|Left|Right|F1|period|Tab|Up|Down|ISO_Left_Tab)", gtk.gdk.keyval_name (event.keyval)):
            return True

    ''' We return True here regardless as we do our matching with the SQL and the default search function is startswith() '''
    def completion_match(self, completion, key, iter, column):
        return True

    def update_autocomplete(self, widget):
        text = widget.get_text()
        model = gtk.ListStore(str)
        
        self.cur.execute("SELECT scientific FROM Taxa WHERE scientific LIKE '%" + text + "%'")

        for record in self.cur:
            model.append([record[0]])
            
        self.cur.execute("SELECT vernacular FROM Taxa WHERE vernacular LIKE '%" + text + "%'")

        for record in self.cur:
            model.append([record[0]])        

        self.cur.execute("SELECT code FROM Taxa WHERE code LIKE '%" + text + "%'")

        for record in self.cur:
            model.append([record[0]])            

        widget.get_completion().set_model(model)

    def on_completion_match(self, completion, model, iter):
        completion.get_entry().set_text(model[iter][0])

    def t2str(self, t):
        array = []

        for item in t:
            item = str(item)
            item = string.replace(item, "<i>", "")
            item = string.replace(item, "</i>", "")
            array.append(item)

        return array

    def statistics_copy (self, widget):
      clipboard = gtk.clipboard_get(gtk.gdk.SELECTION_CLIPBOARD)
      treeselection = self.glade.get_object("treeview1").get_selection()
      (model, paths) = treeselection.get_selected_rows()

      record_list = []

      ncolumns = range(model.get_n_columns())

      for path in paths:                                          
          record_list.append('\t'.join(self.t2str(model.get(model.get_iter(path), *ncolumns))))

      clipboard.set_text('\n'.join(record_list))

    def statistics_select_all (self, widget):
      treeselection = self.glade.get_object("treeview1").get_selection()
      treeselection.select_all()

    def statistics_select_none (self, widget):
      treeselection = self.glade.get_object("treeview1").get_selection()
      treeselection.unselect_all()

    def statistics_window_popup_menu(self, treeview, event):
      if event.button == 3:
          x = int(event.x)
          y = int(event.y)
          time = event.time
          pthinfo = treeview.get_path_at_pos(x, y)
          if pthinfo is not None:
              path, col, cellx, celly = pthinfo
              treeview.grab_focus()

              treeselection = treeview.get_selection()

              if treeselection.count_selected_rows() > 0:
                  self.glade.get_object("menuitem13").set_sensitive(True)
              else:
                  self.glade.get_object("menuitem13").set_sensitive(False)

              self.glade.get_object("statistics_window_popup_menu").popup( None, None, None, event.button, time)
          return True                    

    def del_key (self, widget, event):
    
        if re.match ("(Delete)", gtk.gdk.keyval_name (event.keyval)):
            
            treeselection = self.glade.get_object("treeview2").get_selection()
            (model, paths) = treeselection.get_selected_rows()

            year, month, day = self.glade.get_object("calendar2").get_date()

            for path in paths:
                id = str(model.get_value(model.get_iter(path), 0))
                self.cur.execute("DELETE FROM records WHERE species = '" + id + "' AND date = '" + str(year) + "-" + str(month+1).rjust(2, "0") + "-" + str(day).rjust(2, "0") + "'")
                self.con.commit()
                           
            self.year_change(None)
            self.month_change(None)
            self.day_change(None)                    

    def main_window_delete (self, widget):        
        treeselection = self.glade.get_object("treeview2").get_selection()
        (model, paths) = treeselection.get_selected_rows()

        year, month, day = self.glade.get_object("calendar2").get_date()

        for path in paths:
            id = str(model.get_value(model.get_iter(path), 0))
            self.cur.execute("DELETE FROM records WHERE species = '" + id + "' AND date = '" + str(year) + "-" + str(month+1).rjust(2, "0") + "-" + str(day).rjust(2, "0") + "'")
            self.con.commit()
                           
        self.year_change(None)
        self.month_change(None)
        self.day_change(None)

    def main_window_edit (self, widget):
        treeselection = self.glade.get_object("treeview2").get_selection()
        (model, paths) = treeselection.get_selected_rows()
      
        vernacular = model.get_value(model.get_iter(paths[0]), 1)
        scientific = model.get_value(model.get_iter(paths[0]), 2)
        scientific = string.replace(scientific, "<i>", "")
        scientific = string.replace(scientific, "</i>", "")
            
        count = model.get_value(model.get_iter(paths[0]), 5)
        
        if vernacular == "":
            species = scientific
        else:
            species = scientific
        
        self.glade.get_object("entry6").set_text(species)
        self.glade.get_object("entry4").set_text(str(count))
                               
        self.glade.get_object("label14").set_text("Edit")

    def main_window_copy (self, widget):
      clipboard = gtk.clipboard_get(gtk.gdk.SELECTION_CLIPBOARD)
      treeselection = self.glade.get_object("treeview2").get_selection()
      (model, paths) = treeselection.get_selected_rows()

      record_list = []

      ncolumns = range(model.get_n_columns())

      for path in paths:
          record_list.append('\t'.join(self.t2str(model.get(model.get_iter(path), *ncolumns))))

      clipboard.set_text('\n'.join(record_list))

    def main_window_select_all (self, widget):
      treeselection = self.glade.get_object("treeview2").get_selection()
      treeselection.select_all()

    def main_window_select_none (self, widget):
      treeselection = self.glade.get_object("treeview2").get_selection()
      treeselection.unselect_all()
                                       
    def main_window_popup_menu(self, treeview, event):
      if event.button == 3:
          x = int(event.x)
          y = int(event.y)
          time = event.time
          pthinfo = treeview.get_path_at_pos(x, y)
          if pthinfo is not None:
              path, col, cellx, celly = pthinfo
              treeview.grab_focus()

              treeselection = treeview.get_selection()

              if treeselection.count_selected_rows() > 0:
                  self.glade.get_object("menuitem24").set_sensitive(True)
                  self.glade.get_object("menuitem28").set_sensitive(True)
                  self.glade.get_object("menuitem30").set_sensitive(True)
              else:
                  self.glade.get_object("menuitem24").set_sensitive(False)
                  self.glade.get_object("menuitem28").set_sensitive(False)
                  self.glade.get_object("menuitem30").set_sensitive(False)

              self.glade.get_object("main_window_popup_menu").popup( None, None, None, event.button, time)
          return True

    def sort_ranks(self, model, iter1, iter2, column):

        if model.get_value(iter2, 0) == ":":
            return 0

        var1 = int(model.get_value(iter1, column))
        var2 = int(model.get_value(iter2, column))
 
        if var1 == 0:
            return 1
        elif var2 == 0:
            return 0
        elif var1 < var2:
            return -1
        elif var1 == var2:
            return 0
        elif var1 > var2:
            return 1

    def sort_monthly(self, model, iter1, iter2, column):
        var1 = model.get_value(iter1, column)
        var2 = model.get_value(iter2, column)
 
        if var1 < var2:
            return -1
        elif var1 == var2:
            return 0
        elif var1 > var2:
            return 1

    def sort_date(self, model, iter1, iter2, column):
        var1 = model.get_value(iter1, column)
        var2 = model.get_value(iter2, column)

        if var1 == var2:
            return 0
        elif var1 == "-":
            if model.get_sort_column_id()[1] == gtk.SORT_ASCENDING:
                return 1
            elif model.get_sort_column_id()[1] == gtk.SORT_DESCENDING:
                return -1
        elif var2 == "-":
            if model.get_sort_column_id()[1] == gtk.SORT_ASCENDING:
                return -1
            elif model.get_sort_column_id()[1] == gtk.SORT_DESCENDING:
                return 1
        else:
            var1 = var1.split(DATE_SEPARATOR)
            var2 = var2.split(DATE_SEPARATOR)

            if int(var1[1]) < int(var2[1]):
                return -1
            elif int(var1[1]) == int(var2[1]):
                if int(var1[0]) < int(var2[0]):
                    return -1
                elif int(var1[0]) == int(var2[0]):
                    return 0
                elif int(var1[0]) > int(var2[0]):
                    return 1
            elif int(var1[1]) > int(var2[1]):
                return 1

    def sort_bf(self, model, iter1, iter2):

       if model.get_value(iter2, 0) == ":":
           return 0

       try:
           var1 = int(model.get_value(iter1, 0));
           try:
               var2 = int(model.get_value(iter2, 0));
           except ValueError:
               #do int/string compare
               var2 = re.split('([0-9]+)', model.get_value(iter2, 0))

               if var1 < int(var2[1]):
                   return -1
               elif var1 == int(var2[1]):
                   return -1
               elif var1 > int(var2[1]):
                   return 1
           else:
               #do int/int compare
               if  var1 < var2:
                   return -1
               elif var1 == var2:
                   return 0
               elif var1 > var2:
                   return 1
       except ValueError:
           try: 
               var2 = int(model.get_value(iter2, 0));
           except ValueError:
               #do string/string compare
               var1 = re.split('([0-9]+)', model.get_value(iter1, 0))
               var2 = re.split('([0-9]+)', model.get_value(iter2, 0))

               if  var1[1] < var2[1]:
                   return -1
               elif var1[1] == var2[1]:
                   return cmp(var1[2], var2[2])
               elif var1[1] > var2[1]:
                   return 1
           else:
               #do string/int compare
               var1 = re.split('([0-9]+)', model.get_value(iter1, 0))

               if  int(var1[1]) < var2:
                   return -1
               elif int(var1[1]) == var2:
                   return 1
               elif int(var1[1]) > var2:
                   return 1
 
    def stats_close(self, widget):
       self.glade.get_object("statistics_window").hide()
       return True

    def stats_delete(self, widget, userdata):
        self.glade.get_object("statistics_window").hide() 
        return True

    def plot_graph(self, widget):

        try:

            leftborder = 20
            rightborder = 20
            topborder = 20
            bottomborder = 40

            canvas = self.glade.get_object("eventbox1").get_children()[0]

            root = goocanvas.GroupModel()
            canvas.set_root_item_model(root)
            root = canvas.get_root_item()

            leftbounds, topbounds, rightbounds, bottombounds = canvas.get_bounds()

            treeselection = self.glade.get_object("scrolledwindow1").get_children()[0].get_selection()
            (model, iter) = treeselection.get_selected()

            if model.get_value(iter, 1) == 0:

                treeselection = self.glade.get_object("treeview1").get_selection()
                (model, paths) = treeselection.get_selected_rows()

                columns = range(model.get_n_columns()-3)

                columnwidth = (rightbounds-(leftborder)-(rightborder))/2

                #x axis
                goocanvas.polyline_new_line(root, (topbounds+topborder), (leftbounds+leftborder), (leftbounds+leftborder), (bottombounds-bottomborder))

                #y axis
                goocanvas.polyline_new_line(root, (leftbounds+leftborder), (bottombounds-bottomborder), (rightbounds-rightborder), (bottombounds-bottomborder))
          
                maxvalue = 0

                for column in columns:
                    if model.get_value(model.get_iter(paths[0]), column+3) > maxvalue:
                        maxvalue = model.get_value(model.get_iter(paths[0]), column+3)

                if maxvalue < 10:
                    maxvalue = 10
                elif maxvalue < 100:
                    maxvalue = 100
                elif maxvalue < 1000:
                    maxvalue = 1000
                elif maxvalue < 10000:
                    maxvalue = 10000

                for column in columns:
                    columnvalue = ((float(model.get_value(model.get_iter(paths[0]), column+3))/maxvalue) * (bottombounds-bottomborder)) 

                    print (float(model.get_value(model.get_iter(paths[0]), column+3))/maxvalue)

                    goocanvas.Rect(parent = root,
                                   x = (columnwidth*column)+leftborder,
                                   y = (bottombounds-bottomborder)-columnvalue, 
                                   width = columnwidth,
                                   height = columnvalue,
                                           stroke_color = None,
                                   fill_color = "#5f8fc2",
                                           line_width = 0.0)

        except:    
            logging.debug("goocanvas module not loaded - not responding to graphing request")

    def select_stats_mode(self, treeview):

        while gtk.events_pending():
            gtk.main_iteration()

        self.glade.get_object("statistics_window").window.set_cursor(gtk.gdk.Cursor(gtk.gdk.WATCH))
        self.glade.get_object("vbox777").hide()
        self.glade.get_object("scrolledwindow3").show()
        
        while gtk.events_pending():
            gtk.main_iteration()

        # Get the iter of the selected stats function
        treeselection = treeview.get_selection()
        (model, iter) = treeselection.get_selected()

        if iter:
            # Get unique years from the records table
            self.cur.execute("SELECT DISTINCT STRFTIME('%Y', date) FROM records")

            years = []

            for record in self.cur:
                years.append(record[0])
                
            numyears = len(years)

            # The treeview to populate with data
            treeview = self.glade.get_object("treeview1")

            # Blank the treeview
            for column in treeview.get_columns():
                treeview.remove_column(column)

            treeview.set_model(None)

            # Set the selection to multiple
            selection = treeview.get_selection()
            selection.set_mode(gtk.SELECTION_MULTIPLE)

            # Show the headers
            treeview.set_headers_visible(True)

            # Switch on the selected stats
            if model.get_value(iter, 1) == 0:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Specimen Counts</span>")

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, *[str]*(len(years)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)
                liststore.set_sort_func(0, self.sort_bf)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                # Create the columns for each year and calculate the statistics
                for index, year in enumerate(years):
                    column = gtk.TreeViewColumn(year)
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_ranks, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + year + "' ('code' TEXT, \
                                                                              'count' INTEGER \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + year + "' (code, count) \
                                      SELECT records.species, SUM(records.count) AS count \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                      GROUP BY records.species \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + year + ".count, 0)")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + year + " ON _temp_" + year + ".code = taxa.code ")

                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, year in enumerate(years):
                        if not record[index+3] == 0:
                            liststore.append(record)
                            break

                totaliter = liststore.append([":", "", "", 0, 0])

                for index, year in enumerate(years):
                    self.cur.execute("SELECT SUM(records.count) AS value \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                     ")

                    for record in self.cur:
                        liststore.set_value(totaliter, index+3, "<b>&#8721; " + str(int(record[0])) + "</b>")

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore the treeview
                treeview.set_model(liststore)

            elif model.get_value(iter, 1) == 1:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Species Ranks</span>")

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, *[int]*(len(years)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                # Create the columns for each year and calculate the statistics
                for index, year in enumerate(years):
                    column = gtk.TreeViewColumn(year)
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_ranks, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + year + "' ('code' TEXT, \
                                                                              'count' INTEGER \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + year + "' (code, count) \
                                      SELECT records.species, SUM(records.count) AS count \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                      GROUP BY records.species \
                                      ORDER BY 2 DESC \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + year + ".rowid, 0)")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + year + " ON _temp_" + year + ".code = taxa.code ")

                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, year in enumerate(years):
                        if not record[index+3] == 0:
                            liststore.append(record)
                            break

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore the treeview
                treeview.set_model(liststore)

            elif model.get_value(iter, 1) == 2:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Earliest Dates</span>")

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, *[str]*(len(years)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                date_format = "%d/%m"

                # Create the columns for each year and calculate the statistics
                for index, year in enumerate(years):
                    column = gtk.TreeViewColumn(year)
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_date, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + year + "' ('code' TEXT, \
                                                                              'date' TEXT \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + year + "' (code, date) \
                                      SELECT records.species, STRFTIME('" + date_format + "', MIN(records.date)) AS date \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                      GROUP BY records.species \
                                      ORDER BY 2 DESC \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + year + ".date, '-')")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + year + " ON _temp_" + year + ".code = taxa.code ")
 
                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, year in enumerate(years):
                         if not record[index+3] == "-":
                             liststore.append(record)
                             break

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore to the treeview
                treeview.set_model(liststore)

            elif model.get_value(iter, 1) == 3:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Last Dates</span>")

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, *[str]*(len(years)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                date_format = "%d/%m"

                # Create the columns for each year and calculate the statistics
                for index, year in enumerate(years):
                    column = gtk.TreeViewColumn(year)
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_date, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + year + "' ('code' TEXT, \
                                                                              'date' TEXT \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + year + "' (code, date) \
                                      SELECT records.species, STRFTIME('" + date_format + "', MAX(records.date)) AS date \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                      GROUP BY records.species \
                                      ORDER BY 2 DESC \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + year + ".date, '-')")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + year + " ON _temp_" + year + ".code = taxa.code ")

                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, year in enumerate(years):
                        if not record[index+3] == "-":
                            liststore.append(record)
                            break

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore to the treeview
                treeview.set_model(liststore)

            elif model.get_value(iter, 1) == 4:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Cumulative Monthly Specimen Counts</span>")

                month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

                months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, str, *[str]*(len(months)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                # Create the columns for each month and calculate the statistics
                for index, month in enumerate(months):
                    column = gtk.TreeViewColumn(month_names[index])
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_monthly, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + month + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + month + "' ('code' TEXT, \
                                                                              'count' INTEGER \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + month + "' (code, count) \
                                      SELECT records.species, SUM(records.count) AS count \
                                      FROM records \
                                      WHERE STRFTIME('%m', records.date) = '" + month + "' \
                                      GROUP BY records.species \
                                      ORDER BY 2 DESC \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + month + ".count, 0)")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + month + " ON _temp_" + month + ".code = taxa.code ")
                    
                column = gtk.TreeViewColumn("")
                column.pack_start(cell, True)
                column.set_sort_column_id(index+4)
                column.set_expand(True)
                column.set_attributes(cell, markup=index+4)
                column.set_resizable(True)
                treeview.append_column(column)
                
                self.cur.execute("DROP TABLE IF EXISTS '_temp_sum'")

                self.cur.execute("CREATE TEMP TABLE '_temp_sum' ('code' TEXT, \
                                                                 'count' INTEGER \
                                                                )")

                self.cur.execute("INSERT INTO '_temp_sum' (code, count) \
                                  SELECT records.species, SUM(records.count) AS count \
                                  FROM records \
                                  GROUP BY records.species \
                                  ORDER BY 2 DESC \
                                 ")
 
                sql_join_years.append(" IFNULL(_temp_sum.count, 0)")
                sql_join_leftjoin.append("LEFT JOIN _temp_sum ON _temp_sum.code = taxa.code ")

                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))
                                
                for record in self.cur:
                    for index, month in enumerate(months):
                        if not record[index+3] == 0:
                            piter = liststore.append(record)
                            liststore.set_value(piter, 15, "<b>&#8721; " + liststore.get_value(piter, 15) + "</b>") 
                            break
                                                       
                totaliter = liststore.append([":", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

                for index, month in enumerate(months):                
                    self.cur.execute("SELECT IFNULL(SUM(_temp_" + month + ".count), 0) AS value \
                                      FROM '_temp_" + month + "' \
                                     ")

                    for record in self.cur:
                        liststore.set_value(totaliter, index+3, "<b>&#8721; " + str(int(record[0])) + "</b>")  

                                                
                self.cur.execute("SELECT IFNULL(SUM(records.count), 0) AS value \
                                  FROM records \
                                 ")
                                    
                for record in self.cur:
                    liststore.set_value(totaliter, 15, "<big><span weight='bold' underline='double'>&#8721; " + str(int(record[0])) + "</span></big>")                            

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore to the treeview
                treeview.set_model(liststore)

            elif model.get_value(iter, 1) == 5:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Average Monthly Specimen Counts</span>")

                month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

                months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, str, *[str]*(len(months)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                # Create the columns for each month and calculate the statistics
                for index, month in enumerate(months):
                    column = gtk.TreeViewColumn(month_names[index])
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_monthly, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + month + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + month + "' ('code' TEXT, \
                                                                              'count' INTEGER \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + month + "' (code, count) \
                                      SELECT records.species, CAST(ROUND((SUM(records.count)/ " + str(numyears) + "), 0) AS INTEGER) AS count \
                                      FROM records \
                                      WHERE STRFTIME('%m', records.date) = '" + month + "' \
                                      GROUP BY records.species \
                                      ORDER BY 2 DESC \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + month + ".count, 0)")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + month + " ON _temp_" + month + ".code = taxa.code ")

                column = gtk.TreeViewColumn("")
                column.pack_start(cell, True)
                column.set_sort_column_id(index+4)
                column.set_expand(True)
                column.set_attributes(cell, markup=index+4)
                column.set_resizable(True)
                treeview.append_column(column)
                
                self.cur.execute("DROP TABLE IF EXISTS '_temp_sum'")

                self.cur.execute("CREATE TEMP TABLE '_temp_sum' ('code' TEXT, \
                                                                 'count' INTEGER \
                                                                )")

                self.cur.execute("INSERT INTO '_temp_sum' (code, count) \
                                  SELECT records.species, (SUM(records.count)/" + str(numyears) + ") AS count \
                                  FROM records \
                                  GROUP BY records.species \
                                  ORDER BY 2 DESC \
                                 ")
 
                sql_join_years.append(" IFNULL(_temp_sum.count, 0)")
                sql_join_leftjoin.append("LEFT JOIN _temp_sum ON _temp_sum.code = taxa.code ")
                
                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, month in enumerate(months):
                        if not record[index+3] == 0:
                            piter = liststore.append(record)
                            liststore.set_value(piter, 15, "<b>&#8721; " + liststore.get_value(piter, 15) + "</b>") 
                            break                                                       
                            
                totaliter = liststore.append([":", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

                for index, month in enumerate(months):                
                    self.cur.execute("SELECT IFNULL(SUM(_temp_" + month + ".count), 0) AS value \
                                      FROM '_temp_" + month + "' \
                                     ")

                    for record in self.cur:
                        liststore.set_value(totaliter, index+3, "<b>&#8721; " + str(int(record[0])) + "</b>")  

                self.cur.execute("SELECT (SUM(records.count)/" + str(numyears) + ") AS value \
                                  FROM records \
                                 ")
                                    
                for record in self.cur:
                    liststore.set_value(totaliter, 15, "<big><span weight='bold' underline='double'>&#8721; " + str(int(record[0])) + "</span></big>")                            

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore to the treeview
                treeview.set_model(liststore)

                '''
                    self.cur.execute("SELECT COUNT(DISTINCT records.species) AS value \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                     ")
 
                    for record in self.cur:
                        liststore.set_value(speciesiter, index+1, record[0])

                    self.cur.execute("SELECT SUM(records.count) AS value \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                     ")
 
                    for record in self.cur:
                        liststore.set_value(countiter, index+1, record[0])
                '''
                
            elif model.get_value(iter, 1) == 8:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>" + str(model.get_value(iter, 2)) + " Monthly Specimen Counts</span>")

                month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

                months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, str, *[str]*(len(months)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                # Create the columns for each month and calculate the statistics
                for index, month in enumerate(months):
                    column = gtk.TreeViewColumn(month_names[index])
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_monthly, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + month + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + month + "' ('code' TEXT, \
                                                                              'count' INTEGER \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + month + "' (code, count) \
                                      SELECT records.species, SUM(records.count) AS count \
                                      FROM records \
                                      WHERE STRFTIME('%m', records.date) = '" + month + "' \
                                      AND STRFTIME('%Y', records.date) = '" + str(model.get_value(iter, 2)) + "' \
                                      GROUP BY records.species \
                                      ORDER BY 2 DESC \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + month + ".count, 0)")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + month + " ON _temp_" + month + ".code = taxa.code ")
                
                column = gtk.TreeViewColumn("")
                column.pack_start(cell, True)
                column.set_sort_column_id(index+4)
                column.set_expand(True)
                column.set_attributes(cell, markup=index+4)
                column.set_resizable(True)
                treeview.append_column(column)
                
                self.cur.execute("DROP TABLE IF EXISTS '_temp_sum'")

                self.cur.execute("CREATE TEMP TABLE '_temp_sum' ('code' TEXT, \
                                                                 'count' INTEGER \
                                                                )")

                self.cur.execute("INSERT INTO '_temp_sum' (code, count) \
                                  SELECT records.species, SUM(records.count) AS count \
                                  FROM records \
                                  WHERE STRFTIME('%Y', records.date) = '" + str(model.get_value(iter, 2)) + "' \
                                  GROUP BY records.species \
                                  ORDER BY 2 DESC \
                                 ")
 
                sql_join_years.append(" IFNULL(_temp_sum.count, 0)")
                sql_join_leftjoin.append("LEFT JOIN _temp_sum ON _temp_sum.code = taxa.code ")

                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, month in enumerate(months):
                        if not record[index+3] == 0:
                            piter = liststore.append(record)
                            liststore.set_value(piter, 15, "<b>&#8721; " + liststore.get_value(piter, 15) + "</b>") 
                            break
                            
                totaliter = liststore.append([":", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])

                for index, month in enumerate(months):                
                    self.cur.execute("SELECT IFNULL(SUM(records.count), 0) AS value \
                                      FROM records \
                                      WHERE STRFTIME('%m', records.date) = '" + month + "' \
                                      AND STRFTIME('%Y', records.date) = '" + str(model.get_value(iter, 2)) + "' \
                                     ")

                    for record in self.cur:
                        liststore.set_value(totaliter, index+3, "<b>&#8721; " + str(int(record[0])) + "</b>")
                        
                self.cur.execute("SELECT IFNULL(SUM(records.count), 0) AS value \
                                  FROM records \
                                  WHERE STRFTIME('%Y', records.date) = '" + str(model.get_value(iter, 2)) + "' \
                                 ")
                                    
                for record in self.cur:
                    liststore.set_value(totaliter, 15, "<big><span weight='bold' underline='double'>&#8721; " + str(int(record[0])) + "</span></big>")                            

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + month + "'")

                # Attach the liststore to the treeview
                treeview.set_model(liststore)
                
            elif model.get_value(iter, 1) == 9:
                # Set the title of the stats
                self.glade.get_object("label37").set_markup("<span weight='bold' size='11000' color='white'>Species Counts</span>")

                # Create the liststore
                liststore = gtk.ListStore(str, str, str, *[int]*(len(years)))
                liststore.set_sort_column_id(0, gtk.SORT_ASCENDING)
                liststore.set_sort_func(0, self.sort_bf)

                # Create the columns
                cell = gtk.CellRendererText()

                column = gtk.TreeViewColumn("Code")
                column.pack_start(cell, True)
                column.set_sort_column_id(0)
                column.set_expand(False)
                column.set_attributes(cell, text=0)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Common Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(1)
                column.set_expand(True)
                column.set_attributes(cell, text=1)
                column.set_resizable(True)
                treeview.append_column(column)

                column = gtk.TreeViewColumn("Scientific Name")
                column.pack_start(cell, True)
                column.set_sort_column_id(2)
                column.set_expand(True)
                column.set_attributes(cell, markup=2)
                column.set_resizable(True)
                treeview.append_column(column)

                sql_join_years = []
                sql_join_leftjoin = []
                sql_join_years.append("SELECT taxa.code, taxa.vernacular, taxa.scientific")

                # Create the columns for each year and calculate the statistics
                for index, year in enumerate(years):
                    column = gtk.TreeViewColumn(year)
                    column.pack_start(cell, True)
                    column.set_sort_column_id(index+3)
                    column.set_attributes(cell, markup=index+3)
                    column.set_resizable(True)
                    treeview.append_column(column)
                    liststore.set_sort_func(index+3, self.sort_ranks, index+3)

                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                    self.cur.execute("CREATE TEMP TABLE '_temp_" + year + "' ('code' TEXT, \
                                                                              'count' INTEGER \
                                                                             )")

                    self.cur.execute("INSERT INTO '_temp_" + year + "' (code, count) \
                                      SELECT records.species, SUM(records.count) AS count \
                                      FROM records \
                                      WHERE STRFTIME('%Y', records.date) = '" + year + "' \
                                      GROUP BY records.species \
                                     ")
 
                    sql_join_years.append(" IFNULL(_temp_" + year + ".count, 0)")
                    sql_join_leftjoin.append("LEFT JOIN _temp_" + year + " ON _temp_" + year + ".code = taxa.code ")

                self.cur.execute(' '.join([','.join(sql_join_years), "FROM taxa ", ' '.join(sql_join_leftjoin)]))

                for record in self.cur:
                    for index, year in enumerate(years):
                        if not record[index+3] == 0:
                            liststore.append(record)
                            break

                # Clean up after ourselves
                for index, year in enumerate(years):
                    self.cur.execute("DROP TABLE IF EXISTS '_temp_" + year + "'")

                # Attach the liststore the treeview
                treeview.set_model(liststore)

                
                
                
                
        while gtk.events_pending():
            gtk.main_iteration()
                
        self.glade.get_object("statistics_window").window.set_cursor(None)
        self.glade.get_object("vbox777").show()
        self.glade.get_object("scrolledwindow3").hide()


    def show_statistics(self, widget):

        global stats_window_created

        self.glade.get_object("statistics_window").show()

        if not stats_window_created:
            self.glade.get_object("eventbox6").modify_bg(gtk.STATE_NORMAL, gtk.gdk.color_parse("#5f8fc2"))

            store = gtk.TreeStore(str, int, int)
            treeview = gtk.TreeView(store)

            treeview.set_headers_visible(False)










            iter = store.append(None, ["Specimen Counts", 0, -1])

            self.cur.execute("SELECT DISTINCT STRFTIME('%Y', date) FROM records")

            years = []

            for record in self.cur:
                store.append(iter, [record[0], 8, int(record[0])])

            store.append(iter, ["Cumulative", 4, -1])
            store.append(iter, ["Average", 5, -1])
            treeview.expand_row(store.get_path(iter), False)







            iterp = store.append(None, ["Specimen Ranks", 1, -1])

            iter = store.append(iterp, ["per Month", -1, -1])

            self.cur.execute("SELECT DISTINCT STRFTIME('%Y', date) FROM records")

            years = []

            for record in self.cur:
                store.append(iter, [record[0], -1, int(record[0])])

            store.append(iter, ["Cumulative", -1, -1])
            store.append(iter, ["Average", -1, -1])

            treeview.expand_row(store.get_path(iterp), False)







            iter = store.append(None, ["Phenology", -1, -1])
            store.append(iter, ["Earliest Dates", 2, -1])
            store.append(iter, ["Lastest Dates", 3, -1])

            treeview.expand_row(store.get_path(iter), False)

            cell = gtk.CellRendererText()
            col = gtk.TreeViewColumn("Statistics")
            col.pack_start(cell, True)
            col.set_attributes(cell, markup=0)

            treeview.append_column(col)

            self.glade.get_object("scrolledwindow1").add(treeview)
            self.glade.get_object("scrolledwindow1").show_all()

            treeview.connect("cursor_changed", self.select_stats_mode)

            stats_window_created = True

            treeselection = treeview.get_selection()

            try: 
                canvas = goocanvas.Canvas()
            except:
                logging.debug("goocanvas module not loaded - not creating canvas") 
                
            else:            
                bg_color = gtk.gdk.Color (65535, 65535, 65535, 0)
                canvas.modify_base (gtk.STATE_NORMAL, bg_color)
                canvas.set_size_request(-1, 200)
                canvas.set_bounds(0, 0, 500, 200)

                canvas.show()

                root = goocanvas.GroupModel()
                canvas.set_root_item_model(root)
            
                self.glade.get_object("eventbox1").add(canvas)
                self.glade.get_object("expander1").show()

            while gtk.events_pending():
                gtk.main_iteration()

            self.select_stats_mode(treeview)

    def begin_print(self, operation, context, print_data):
        width = context.get_width()
        height = context.get_height()
        print_data.layout = context.create_pango_layout()
        print_data.layout.set_font_description(pango.FontDescription("Monospace 12"))
        print_data.layout.set_width(int(width*pango.SCALE))
        print_data.layout.set_text(print_data.text)

        print_data.layout.set_markup (print_data.text)

        num_lines = print_data.layout.get_line_count()

        page_breaks = []
        page_height = 0

        for line in xrange(num_lines):
            layout_line = print_data.layout.get_line(line)
            ink_rect, logical_rect = layout_line.get_extents()
            lx, ly, lwidth, lheight = logical_rect
            line_height = lheight / 1024.0

            if page_height + line_height > height:
                page_breaks.append(line)
                page_height = 0

            page_height += line_height

        operation.set_n_pages(len(page_breaks) + 1)
        print_data.page_breaks = page_breaks


    def draw_page(self, operation, context, page_nr, print_data):
        assert isinstance(print_data.page_breaks, list)
        if page_nr == 0:
            start = 0
        else:
            start = print_data.page_breaks[page_nr - 1]

        try:
            end = print_data.page_breaks[page_nr]
        except IndexError:
            end = print_data.layout.get_line_count()

        cr = context.get_cairo_context()

        cr.set_source_rgb(0, 0, 0)

        i = 0
        start_pos = 0
        iter = print_data.layout.get_iter()
        while 1:
            if i >= start:
                line = iter.get_line()
                _, logical_rect = iter.get_line_extents()
                lx, ly, lwidth, lheight = logical_rect
                baseline = iter.get_baseline()
                if i == start:
                    start_pos = ly / 1024.0;
                cr.move_to(lx / 1024.0, baseline / 1024.0 - start_pos)
                cr.show_layout_line(line)
            i += 1
            if not (i < end and iter.next_line()):
                break


    def do_page_setup(self, action):
        global settings, page_setup
        if settings is None:
            settings = gtk.PrintSettings()
        page_setup = gtk.print_run_page_setup_dialog(self.glade.get_object("main_window"),
                                                        page_setup, settings)

    def do_print(self, action):
        global settings, page_setup
        print_data = PrintData()

        print_data.text = "<small>Sandby Moth Recorder " + gobject.markup_escape_text(self.version) + "</small>\n\n"
        print_data.text += "<b><big>" + gobject.markup_escape_text(self.glade.get_object("entry1").get_text()) + " - " + gobject.markup_escape_text(self.glade.get_object("entry3").get_text()) + "</big></b>\n"
        print_data.text += gobject.markup_escape_text(self.glade.get_object("entry2").get_text()) + "\n"

        self.cur.execute("SELECT * FROM records ORDER BY date, species + 0.0")

        date = False

        for rcd in self.cur:
            if date != rcd[0]:
                year = rcd[0][0:4]
                month = rcd[0][5:7]
                day = rcd[0][8:10]

                print_data.text += "\n<b>" + gobject.markup_escape_text(day + "/" + month + "/" + year) + "</b>\n"

            date = rcd[0]

            if (rcd[1] == "0"):
                print_data.text += "No catch\n"
            else:
                print_data.text += gobject.markup_escape_text(str(rcd[1]).ljust(5, " ")) + " " + gobject.markup_escape_text(self.sandby_dictionary[str(rcd[1])][0].ljust(40, " ")) + " <i>" + gobject.markup_escape_text(self.sandby_dictionary[str(rcd[1])][1].ljust(25, " ")) + "</i> " + " " + gobject.markup_escape_text(self.sandby_dictionary[str(rcd[1])][3].ljust(15, " ")) + " " + gobject.markup_escape_text(str(rcd[2]).ljust(4, " ")) + "\n"

        print_ = gtk.PrintOperation()
        if settings is not None:
            print_.set_print_settings(settings)

        if page_setup is not None:
            print_.set_default_page_setup(page_setup)

        print_.connect("begin_print", self.begin_print, print_data)
        print_.connect("draw_page", self.draw_page, print_data)

        try:
            res = print_.run(gtk.PRINT_OPERATION_ACTION_PRINT_DIALOG, self.glade.get_object("main_window"))
        except gobject.GError, ex:
            error_dialog = gtk.MessageDialog(self.glade.get_object("main_window"),
                                             gtk.DIALOG_DESTROY_WITH_PARENT,
                                             gtk._MESSAGE_ERROR,
                                             gtk.BUTTONS_CLOSE,
                                             ("Error printing file:\n%s" % str(ex)))
            error_dialog.connect("response", gtk.Widget.destroy)
            error_dialog.show()
        else:
            if res == gtk.PRINT_OPERATION_RESULT_APPLY:
                settings = print_.get_print_settings()

        if not print_.is_finished():
            active_prints.remove(print_)

    def window_resize(self, widget, userdata):
       if self.glade.get_object("togglebutton1").get_active():
         """Called when the due button is clicked."""
         rect = self.glade.get_object("togglebutton1").get_allocation()
         x, y = self.glade.get_object("togglebutton1").window.get_origin()

         self.glade.get_object("calendar_window").show()
         self.glade.get_object("calendar_window").move((x + rect.x), (y + rect.y + rect.height))

    def toggle_calendar(self, widget):
       """Called when the due button is clicked."""
       if widget.get_active():
           rect = widget.get_allocation()
           x, y = widget.window.get_origin()

           self.glade.get_object("calendar_window").show() 
           self.glade.get_object("calendar_window").move((x + rect.x), (y + rect.y + rect.height))
       else:
           self.glade.get_object("calendar_window").hide()

    def quit(self, widget):
        gtk.main_quit()
        sys.exit()

    def dataset_properties(self, widget):
        self.glade.get_object("dataset_properties").show()
        self.glade.get_object("notebook1").set_current_page(0)

    def dataset_properties_close(self, widget):
        name = self.glade.get_object("entry1").get_text()
        description = self.glade.get_object("entry2").get_text()
        author = self.glade.get_object("entry3").get_text()

        self.cur.execute("UPDATE meta SET name = '" + name + "', description = '" + description + "', author = '" + author + "'")
        self.con.commit()
        self.glade.get_object("dataset_properties").hide()
        self.load_metadata()
        return True

    def dataset_properties_hide(self, widget, userdata):
        self.glade.get_object("dataset_properties").hide()
        return True

    def show_msg(self, primary_text, seconday_text, icon="error", buttons="close"):
        """Show message dialog"""
        msgicon=0
        msgbuttons=0
        if icon=="error":
            msgicon=gtk.MESSAGE_ERROR
        elif icon=="info":
            msgicon=gtk.MESSAGE_INFO
        elif icon=="question":
            msgicon=gtk.MESSAGE_QUESTION
        elif icon=="warning":
            msgicon=gtk.MESSAGE_WARNING

        if buttons=="close":
            msgbuttons=gtk.BUTTONS_CLOSE
        elif buttons=="yesno":
            msgbuttons=gtk.BUTTONS_YES_NO
        elif buttons=="ok":
            msgbuttons=gtk.BUTTONS_OK

        msgWindow=gtk.MessageDialog(None, 0, msgicon, msgbuttons, primary_text)
        msgWindow.format_secondary_text(seconday_text)
        response=msgWindow.run()
        msgWindow.destroy()
        return response

    def set_widgets_sensitive(self,state):
        """change sensitive object after opening/closing database"""
        #self.glade.get_object("menuitem5").set_sensitive(state)
        #self.glade.get_object("button6").set_sensitive(state)
        #self.glade.get_object("button7").set_sensitive(state)
        #self.glade.get_object("button11").set_sensitive(state)
        #self.glade.get_object("button12").set_sensitive(state)
        #self.glade.get_object("imagemenuitem4").set_sensitive(state)
        #self.glade.get_object("tool_add_row").set_sensitive(state)
        #self.glade.get_object("imagemenuitem6").set_sensitive(state)
        #self.glade.get_object("menuitem6").set_sensitive(state)
        #self.glade.get_object("menuitem7").set_sensitive(state)

    def create_database(self, filename, checklist):
        """create database file"""
        try :
            file_handler=open(filename,'w+')
            self.database_filename=filename
            
            if os.name == 'nt':
                filenameprint = filename
            else:
                filenameprint = filename.replace (os.getenv("HOME"), '~', 1)
   
            shutil.copy2(checklists_data[checklist][1], filename)

            self.set_widgets_sensitive(True)
            self.con = sqlite.connect(self.database_filename)
            self.cur = self.con.cursor()
            self.load_tables()
        except IOError:
            self.show_msg("An IO error occured.", "", "error", "close")
            return

    def copy_database(self,filename):
        """copy (Save as database)"""
        try :            
            shutil.copy2(self.database_filename, filename)
            self.database_filename=filename
            self.con = None
            self.cur = None
            self.load_tables()
        except IOError:
            self.show_msg("An IO error occured.", "", "error", "close")
            return


    def open_database(self,widget):
        dialog = gtk.FileChooserDialog("Open Dataset",
                                       None,
                                       gtk.FILE_CHOOSER_ACTION_OPEN,
                                       (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                        gtk.STOCK_OPEN, gtk.RESPONSE_OK))
        dialog.set_default_response(gtk.RESPONSE_OK)

        dialog.set_transient_for(self.glade.get_object("main_window"))
        dialog.set_property('skip-taskbar-hint', True)

        response = dialog.run()
        self.database_filename=dialog.get_filename()
        dialog.destroy()

        if response == gtk.RESPONSE_OK:
            self.load_tables()




    def new_database(self,widget):
       dialog = gtk.FileChooserDialog("New Dataset ",
                                      None,
                                      gtk.FILE_CHOOSER_ACTION_SAVE,
                                      (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                       gtk.STOCK_SAVE, gtk.RESPONSE_OK)
                                     )

       dialog.set_default_response(gtk.RESPONSE_OK)
       dialog.set_position(gtk.WIN_POS_CENTER_ON_PARENT)
       dialog.set_transient_for(self.glade.get_object("main_window"))
       dialog.set_property('skip-taskbar-hint', True)

       hbox = gtk.HBox(False, 10)

       combobox = gtk.combo_box_new_text()

       for clist in checklists_data:
           combobox.append_text(clist[0])

       combobox.set_active(0)

       label = gtk.Label("Using checklist:")

       hbox.pack_start(label, False, True, 0)
       hbox.pack_start(combobox, True, True, 0)

       hbox.show_all()

       dialog.set_extra_widget(hbox)

       response = dialog.run()
       if response == gtk.RESPONSE_OK:
           filename=dialog.get_filename()
           clist = dialog.get_extra_widget().get_children()[1].get_active()

           dialog.destroy()
           if os.path.exists(filename) ==True:

               if self.show_msg("A file named \"" + os.path.basename(filename) + "\" already exists.\nDo you want to replace it?", "The file already exists in \"" + os.path.dirname(filename) + "\". Replacing it will overwrite its contents.", "question", "yesno")==gtk.RESPONSE_YES:
                   self.create_database(filename, clist)

           else:
               self.create_database(filename, clist)

       else:
           dialog.destroy()

    def close_database(self,widget):
        """Close Database"""
        if widget!="":
            self.main_window.set_title('Sandby Moth Recorder')
            self.set_widgets_sensitive(False) # >>> v

            #self.glade.get_object("hbox1").set_sensitive(False)
            self.glade.get_object("imagemenuitem4").set_sensitive(False)
            self.glade.get_object("menuitem6").set_sensitive(False)
            self.glade.get_object("menuitem8").set_sensitive(False)
            self.glade.get_object("menuitem2").set_sensitive(False)
            self.glade.get_object("menuitem15").set_sensitive(False)
            self.glade.get_object("menuitem7").set_sensitive(False)
            self.glade.get_object("menuitem3").set_sensitive(False)
            self.glade.get_object("menuitem5").set_sensitive(False)
            self.glade.get_object("toolbutton2").set_sensitive(False)
            self.glade.get_object("togglebutton1").hide()
            self.glade.get_object("hbox3").hide()
            self.database_table.set_headers_visible(False)

            self.database_table_rows.clear()

            self.con=None
            self.cur=None
            self.database_filename=None

    def save_database_as(self,widget):
       """save as new database"""
       dialog = gtk.FileChooserDialog("Save As",
                                    None,
                                    gtk.FILE_CHOOSER_ACTION_SAVE,
                                    (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                     gtk.STOCK_SAVE, gtk.RESPONSE_OK))
       dialog.set_default_response(gtk.RESPONSE_OK)
       response = dialog.run()

       dialog.set_transient_for(self.glade.get_object("main_window"))
       dialog.set_property('skip-taskbar-hint', True)

       if response == gtk.RESPONSE_OK:
           filename=dialog.get_filename()
           dialog.destroy()
           if os.path.exists(filename) ==True:
               if self.show_msg("A file named \"" + os.path.basename(filename) + "\" already exists.\nDo you want to replace it?", "The file already exists in \"" + os.path.dirname(filename) + "\". Replacing it will overwrite its contents.", "question", "yesno")==gtk.RESPONSE_YES:

                   self.copy_database(filename)

           else:

               self.copy_database(filename)
               pass

       else:
           dialog.destroy()

    def load_tables(self):
        """load databases and fill table in treeview"""
        try:

            self.con = sqlite.connect(self.database_filename)
            self.cur = self.con.cursor()
            
            self.load_metadata()

            self.set_widgets_sensitive(True)

            today = date.today()

            self.glade.get_object("calendar2").select_month(today.month-1, today.year)
            self.glade.get_object("calendar2").select_day(today.day)

            self.year_change(False)
            self.month_change(False)
            self.day_change(False)

            #self.glade.get_object("hbox1").set_sensitive(True)
            self.glade.get_object("imagemenuitem4").set_sensitive(True)
            self.glade.get_object("menuitem6").set_sensitive(True)
            self.glade.get_object("menuitem8").set_sensitive(True)
            self.glade.get_object("menuitem2").set_sensitive(True)
            self.glade.get_object("menuitem15").set_sensitive(True)
            self.glade.get_object("menuitem7").set_sensitive(True)
            self.glade.get_object("menuitem3").set_sensitive(True)
            self.glade.get_object("menuitem5").set_sensitive(True)
            self.glade.get_object("toolbutton2").set_sensitive(True)

            self.database_table.set_headers_visible(True)


            self.glade.get_object("calendar_window").realize()
            rect = self.glade.get_object("calendar_window").get_allocation()


            self.glade.get_object("togglebutton1").set_size_request(rect.width-1, -1)

            self.glade.get_object("togglebutton1").show()
            self.glade.get_object("hbox3").show()

            treeview = self.glade.get_object("scrolledwindow6").get_child()
            treeview.set_headers_visible(True)
        except sqlite.DatabaseError:
            self.show_msg("Could not open the file " + self.database_filename + ".", "The file does not appear to be a valid dataset.", "error", "close")
            self.con=None
            self.cur=None
            self.database_filename=None
            return

    def add_data(self,widget):
        year, month, day = self.glade.get_object("calendar2").get_date()
        count = int(self.glade.get_object("entry4").get_text())
        entry = self.glade.get_object("entry6").get_text()

        if count != 0:

          self.cur.execute("SELECT code FROM taxa WHERE scientific LIKE ? OR vernacular LIKE ? OR code LIKE ?", [entry, entry, entry])

          species = None

          for record in self.cur:
              species = record[0]                      
        
          if species:
              self.cur.execute("INSERT INTO records (date, count, species) VALUES('" + str(year) + "-" + str(month+1).rjust(2, "0") + "-" + str(day).rjust(2, "0") + "', '" + str(count) + "', '" + species + "')")
              self.con.commit()

              self.glade.get_object("entry6").set_text("Species")
              self.glade.get_object("entry4").set_text("Count")

              self.year_change(None)
              self.month_change(None)
              self.day_change(None, True)
          else:
              dialog = gtk.MessageDialog(self.glade.get_object("main_window"), 
                                         gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_WARNING, 
                                         gtk.BUTTONS_OK)
              dialog.set_position(gtk.WIN_POS_CENTER_ON_PARENT)
              dialog.set_markup("<big><b>Taxon doesn't exist</b></big>")
              dialog.format_secondary_text("'" + entry + "' is not in the taxon dictionary.")
              dialog.run()
              dialog.destroy()

    def do_data(self, widget):
        if self.glade.get_object("label14").get_text() == "Add":
            self.add_data(widget)
        if self.glade.get_object("label14").get_text() == "Edit":
            self.edit_data(widget)

    def edit_data(self, widget):
        year, month, day = self.glade.get_object("calendar2").get_date()
        count = int(self.glade.get_object("entry4").get_text())
        entry = self.glade.get_object("entry6").get_text()
        
        if count != 0:

          self.cur.execute("SELECT code FROM taxa WHERE scientific LIKE ? OR vernacular LIKE ? OR code LIKE ?", [entry, entry, entry])

          species = None

          for record in self.cur:
              species = record[0]                      
        
          if species:
              treeselection = self.glade.get_object("treeview2").get_selection()
              (model, paths) = treeselection.get_selected_rows()
      
              self.cur.execute("SELECT code FROM taxa WHERE code LIKE ?", [model.get_value(model.get_iter(paths[0]), 0)])
            
              for record in self.cur:
                  oldspecies = record[0] 
                  
              self.cur.execute("UPDATE records SET species='" + species + "', count='" + str(count) + "' WHERE species='" + oldspecies + "' AND date='" + str(year) + "-" + str(month+1).rjust(2, "0") + "-" + str(day).rjust(2, "0") + "'")
              self.con.commit()

              self.glade.get_object("entry6").set_text("Species")
              self.glade.get_object("entry4").set_text("Count")                          
      
              self.glade.get_object("label14").set_text("Add")
        
              self.year_change(None)
              self.month_change(None)
              self.day_change(None)
          else:
              dialog = gtk.MessageDialog(self.glade.get_object("main_window"), 
                                         gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_WARNING, 
                                         gtk.BUTTONS_OK)
              dialog.set_position(gtk.WIN_POS_CENTER_ON_PARENT)
              dialog.set_markup("<big><b>Taxon doesn't exist</b></big>")
              dialog.format_secondary_text("'" + entry + "' is not in the taxon dictionary.")
              dialog.run()
              dialog.destroy()

    def year_change(self,widget):
        # attempt to records from the database; handle any exceptions
        try:

            self.glade.get_object("calendar2").clear_marks()

            year, month, day = self.glade.get_object("calendar2").get_date()
            date = str(year)

            day_ord = str(day) + {1: 'st', 2: 'nd', 3: 'rd'}.get(day % (10 < day % 100 < 14 or 10), 'th')

            self.glade.get_object("label3").set_markup(day_ord + " " + calendar.month_name[month+1] + " " + str(year))

        # sqlite error handler
        except sqlite.OperationalError, errormsg:
            self.show_msg("A SQL error has occured.", "The error is: " + errormsg.__str__(), "error", "close")

    def month_change(self,widget):

        year, month, day = self.glade.get_object("calendar2").get_date()

        if month == 0:
            self.year_change(None)

        if month == 11:
            self.year_change(None)

        # attempt to records from the database; handle any exceptions
        try:

            self.glade.get_object("calendar2").clear_marks()

            year, month, day = self.glade.get_object("calendar2").get_date()
            date = str(year) + "-" + str(month+1).rjust(2, "0")

            day_ord = str(day) + {1: 'st', 2: 'nd', 3: 'rd'}.get(day % (10 < day % 100 < 14 or 10), 'th')

            self.glade.get_object("label3").set_markup(day_ord + " " + calendar.month_name[month+1] + " " + str(year))

            self.cur.execute("SELECT * FROM records WHERE date LIKE '%" + date + "%'")

            for rcd in self.cur:
                day = int(rcd[0][8:])
                self.glade.get_object("calendar2").mark_day(day)

        # sqlite error handler
        except sqlite.OperationalError, errormsg:
            self.show_msg("A SQL error has occured.", "The error is: " + errormsg.__str__(), "error", "close")

    """  """
    def day_change_double_click(self,widget):
      self.day_change(widget);
      self.glade.get_object("togglebutton1").set_active(False);
        
    """ Load the records from the database """
    def day_change(self, widget, highlight=False):

        # attempt to load records from the database
        try:

            year, month, day = self.glade.get_object("calendar2").get_date()
            date = str(year) + "-" + str(month+1).rjust(2, "0") + "-" + str(day).rjust(2, "0")

            day_ord = str(day) + {1: 'st', 2: 'nd', 3: 'rd'}.get(day % (10 < day % 100 < 14 or 10), 'th')

            self.glade.get_object("label3").set_markup(day_ord + " " + calendar.month_name[month+1] + " " + str(year))

            self.database_table_rows.clear()

            self.cur.execute("SELECT records.species, records.count, taxa.scientific, taxa.vernacular, taxa.authority, taxa.family FROM records, taxa WHERE records.date LIKE '" + date + "' AND records.species = taxa.code")

            count = 0

            for rcd in self.cur:
                count += 1

                if rcd[1] != "0":
                    iter=self.database_table_rows.append([rcd[0], rcd[3], "<i>" + rcd[2] + "</i>", rcd[4], rcd[5], int(rcd[1])])
            
            if highlight:     
                treeselection = self.glade.get_object("treeview2").get_selection()
                treeselection.select_iter(iter)

            if count > 0:
                self.glade.get_object("label3").set_markup("<b>" + day_ord + " " + calendar.month_name[month+1] + " " + str(year) + "</b>")

            self.set_widgets_sensitive(True)

        # sqlite error handler
        except sqlite.OperationalError, errormsg:
            self.show_msg("A SQL error has occured.", "The error is: " + errormsg.__str__(), "error", "close")

    def import_csv(self,widget):

        dialog = gtk.FileChooserDialog("Import CSV into Dataset",
                                       None,
                                       gtk.FILE_CHOOSER_ACTION_OPEN,
                                       (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                        gtk.STOCK_OPEN, gtk.RESPONSE_OK))
        dialog.set_default_response(gtk.RESPONSE_OK)
        response = dialog.run()

        dialog.set_transient_for(self.glade.get_object("main_window"))
        dialog.set_property('skip-taskbar-hint', True)

        if response == gtk.RESPONSE_OK:
            filename=dialog.get_filename()
            dialog.destroy()

            try:
                reader = csv.reader(open(filename, "rb"))
                insertcount = 0

                reader.next()
                rowcount = 0;

                species_codes = {}
                
                #grab the species codes
                self.cur.execute("SELECT taxa.vernacular, taxa.code \
                                  FROM taxa")

                for rcd in self.cur:
                   species_codes[rcd[0]] = rcd[1]
                
                #grab the species codes
                self.cur.execute("SELECT taxa.scientific, taxa.code \
                                  FROM taxa")

                for rcd in self.cur:
                   species_codes[rcd[0]] = rcd[1]

                for row in reader:

                    date = row[0]
                    species = row[1]
                    count = row[2]
   
                    rowcount += 1

                    try:
                        sql="INSERT INTO records VALUES ('" + date + "', '" + species_codes[species] + "', '" + count + "');"
                        self.cur.execute(sql)
                    except:
                        self.show_msg("Import Error", "Row " + str(rowcount) + ": " + ', '.join(row), "error", "close")

                self.con.commit()
                self.year_change(False)
                self.month_change(False)
                self.day_change(False)

            except:
                self.show_msg("A CSV import error has occured.", "", "error", "close")
            else:
                self.show_msg("Import Complete", filename + " has been imported succesfully.", "info", "ok")

    def export_csv(self,widget):
       dialog = gtk.FileChooserDialog("Export Dataset as CSV",
                                    None,
                                    gtk.FILE_CHOOSER_ACTION_SAVE,
                                    (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                     gtk.STOCK_SAVE, gtk.RESPONSE_OK))
       dialog.set_default_response(gtk.RESPONSE_OK)

       dialog.set_transient_for(self.glade.get_object("main_window"))
       dialog.set_property('skip-taskbar-hint', True)

       dialog.set_current_name(os.path.basename(self.database_filename) + ".csv")
       response = dialog.run()

       if response == gtk.RESPONSE_OK:
           filename=dialog.get_filename()
           dialog.destroy()
           if os.path.exists(filename) == True:
               if self.show_msg("A file named \"" + os.path.basename(filename) + "\" already exists.\nDo you want to replace it?", "The file already exists in \"" + os.path.dirname(filename) + "\". Replacing it will overwrite its contents.", "question", "yesno")==gtk.RESPONSE_YES:

                   writer = csv.writer(open(filename, "wb"))
                   writer.writerow(["date", "bf", "common name", "scientific name", "authority", "family", "count"])

                   writer = csv.writer(open(filename, "wb"))
                   writer.writerow(["date", "code", "common name", "scientific name", "count"])

                   self.cur.execute("SELECT records.date, taxa.code, taxa.vernacular, taxa.scientific, records.count \
                                     FROM records, taxa \
                                     WHERE records.species = taxa.code \
                                     ORDER BY date, code + 0.0 \
                                    ")

                   for rcd in self.cur:
                       writer.writerow(rcd)

           else:

               writer = csv.writer(open(filename, "wb"))
               writer.writerow(["date", "code", "common name", "scientific name", "count"])

               self.cur.execute("SELECT records.date, taxa.code, taxa.vernacular, taxa.scientific, records.count \
                                 FROM records, taxa \
                                 WHERE records.species = taxa.code \
                                 ORDER BY date, code + 0.0 \
                                ")

               for rcd in self.cur:
                   writer.writerow(rcd)

       else:
           dialog.destroy()

    def load_metadata(self):
       self.cur.execute("Select * FROM 'meta'")
       row = self.cur.fetchone()

       self.glade.get_object("entry1").set_text(row[1])
       self.glade.get_object("entry2").set_text(row[2])
       self.glade.get_object("entry3").set_text(row[3])
       self.glade.get_object("label5").set_text(row[0])
       
       if os.name == 'nt':
           filenameprint = self.database_filename
       else:
           filenameprint = self.database_filename.replace (os.getenv("HOME"), '~', 1)

       self.main_window.set_title(row[1] + " (" +  filenameprint + ")" + ' - Sandby Moth Recorder')

       self.glade.get_object("label16").set_text(str(os.path.basename(self.database_filename)))
       self.glade.get_object("label17").set_text(str(os.path.getsize(self.database_filename)/1024) + "KB (" + str(os.path.getsize(self.database_filename)) + " bytes)")

       self.glade.get_object("label26").set_text(str(os.path.dirname(self.database_filename)))
       self.glade.get_object("label27").set_text(time.ctime(os.path.getmtime(self.database_filename)))
       self.glade.get_object("label13").set_text(time.ctime(os.path.getatime(self.database_filename)))

     # track the fullsceen state
    def on_window_state_change(self, widget, event, *args):
        if event.new_window_state & gtk.gdk.WINDOW_STATE_FULLSCREEN:
            self.window_in_fullscreen = True
        else:
            self.window_in_fullscreen = False

     # fullscreen mode switchero
    def switch_fullscreen(self, widget):
        if self.window_in_fullscreen:
            label = self.glade.get_object("menuitem12").get_child()
            label.set_text("Fullscreen")
            img = gtk.image_new_from_stock(gtk.STOCK_FULLSCREEN, gtk.ICON_SIZE_MENU)
            self.glade.get_object("menuitem12").set_image(img)
            self.glade.get_object("main_window").unfullscreen ()
        else:
            label = self.glade.get_object("menuitem12").get_child()
            label.set_text("Leave Fullscreen")
            img = gtk.image_new_from_stock(gtk.STOCK_LEAVE_FULLSCREEN, gtk.ICON_SIZE_MENU)
            self.glade.get_object("menuitem12").set_image(img)
            self.glade.get_object("main_window").fullscreen ()


     # toolbar switchero
    def switch_toolbar(self, widget):
        if self.glade.get_object("menuitem11").get_active() == True:
            self.glade.get_object("toolbar1").show()
        else:
            if self.glade.get_object("togglebutton1").get_active():
                self.glade.get_object("togglebutton1").set_active(False)
                self.glade.get_object("calendar_window").hide()

            self.glade.get_object("toolbar1").hide()

     # statusbar switchero
    def switch_statusbar(self, widget):
        if self.glade.get_object("menuitem14").get_active() == True:
            self.glade.get_object("statusbar1").show()
        else:
            self.glade.get_object("statusbar1").hide()

     # open up the browser and send it to our address
    def on_url(d, link, data):
        if os.name == 'mac':
            subprocess.call(('open', link))
        elif os.name == 'nt':
            subprocess.call(('start', link))
        else:
            subprocess.call(('xdg-open', link))

    gtk.about_dialog_set_url_hook(on_url, None)

     # display the about dialog
    def about(self,widget):
       icon = gtk.gdk.pixbuf_new_from_file("logo.png")

       about=gtk.AboutDialog()
       about.set_name("Sandby\nMoth Recorder")
       about.set_version("")
       about.set_logo(icon)
       about.set_comments(self.version)
       about.set_copyright("Copyright © 2008-2009 Charlie Barnes")
       about.set_authors(["Charlie Barnes <charlie@cucaera.co.uk>"])
       about.set_license("Sandby Moth Recorder is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the Licence, or (at your option) any later version.\n\nSandby Moth Recorder is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with Sandby Moth Recorder; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA")
       about.set_wrap_license(True)
       about.set_website("http://cucaera.co.uk/software/sandby/")
       about.set_transient_for(self.main_window)
       result=about.run()
       about.destroy()

if __name__ == '__main__':
    sandbyActions()
    gtk.main()

