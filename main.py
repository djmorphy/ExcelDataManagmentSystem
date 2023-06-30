from tkinter import*
import tkinter as tk
import tkinter.messagebox as messagebox
from tkinter import ttk
import pandas as pd             #pid install pandas
import matplotlib.pyplot as plt        #pip install matplotlib
import pip
pip.main(["install", "openpyxl"])




class RescueDB:
    def __init__(self,root):
        self.root = root
        self.root.title("Data Managment Syytems")
        self.root.geometry("1920x900+0+0")

        TitleFrame = Frame(self.root, bd = 14, width=1920, height=150, padx=12, relief=RIDGE)
        TitleFrame.grid(row=0, column=0)
        MainFrame = Frame(self.root)
        MainFrame.grid(row=1, column=0)

        TopFrame = Frame(MainFrame, bd=14, width=1350, height=550, padx=4, relief=RIDGE)
        TopFrame.grid(row=0, column=0)

        LeftFrameMain = Frame(TopFrame, bd=10, width=400, height=750, padx=12, relief=RIDGE)
        LeftFrameMain.grid(row=0, column=0)
        LeftFrame = Frame(LeftFrameMain, bd=10, width=450, height=500, relief=RIDGE)
        LeftFrame.grid(row=0, column=0)
        Leftbottom = Frame(LeftFrameMain, bd=10, width=650, height=190, relief=RIDGE)
        Leftbottom.grid(row=1, column=0, pady=1)

        RightFrame = Frame(TopFrame, bd=10, width=1200, height=550, pady=6, relief=RIDGE)
        RightFrame.grid(row=0, column=1)

        BottomFrame = Frame(MainFrame, bd=10, width=1350, height=150, padx=14, relief=RIDGE)
        BottomFrame.grid(row=1, column=0)

    #Function

        def update_data():
            try:
                df = pd.read_excel('Rescue_Dogs.xlsx')
                dog_id = dog_id_entry.get()
                new_data = {
                             'Dog_ID':[dog_id],
                             'Dogs_Name':[dog_name_entry.get()],
                             'Breed':[breed_label_entry.get()],
                             'Colour':[colour_label_entry.get()],
                             'Sex':[sex_label_entry.get()],
                             'Year_of_birth':[year_of_birth_label.get()],
                             'Number of Dogs':[number_of_dogs_entry.get()]
                            }
                new_df = pd.DataFrame(new_data)
                df = pd.concat([df, new_df], ignore_index=True)
                df.to_excel("Rescue_Dogs.xlsx", index = False)
                messagebox.showinfo("Success","Data updated succesfully")


                #clear the entry widget
                reset_entries()

                #Refresh treeview
                refresh_treeview()




            except Exception as e:
                messagebox.showerror("Error",str(e))

            def reset_entries():
                dog_id_entry.delete(0, END)
                dog_name_entry.delete(0, END)
                breed_label_entry.delete(0, END)
                colour_label_entry.delete(0, END)
                sex_label_entry.delete(0, END)
                year_of_birth_entry.delete(0, END)
                number_of_dogs_entry.delete(0, END)

            def refresh_treeview():
                try:
                    df = pd.read_excel('Rescue_Dogs.xlsx')
                    treeview.delete(*treeview.get_children())
                    for index, row in df.iterrows():
                        treeview.insert('', 'end', values=(
                            row['Dog_ID'],
                            row['Dog_Name'],
                            row['Breed'],
                            row['Colour'],
                            row['Sex'],
                            row['Year_of_Birth'],
                            row['Number_of_Dogs']))
                except Exception as e:
                    messagebox.showerror("Error", str(e))

            def exit_program()
                result = messagebox.askquestion('Confirm Exit', 'Are you sure you want to exit?')
                if result == 'yes':
                    root.destroy()
        #Create Title widgets
        dataTitle = Label(TitleFrame, font=('arial', 90, 'bold'), padx=16, text='Excel Data Managment System')
        dataTitle.grid(row=0, column=0)

        subTitle = Label(Leftbottom, font=('arial', 90, 'bold'), padx=16, text='Excel Data')
        subTitle.grid(row=0, column=0)

        #Create Entry widgets

        dog_id_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Dog ID:')
        dog_id_label.grid(row=0, column=0)
        dog_id_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        dog_id_entry.grid(row=0, column=1)

        dog_name_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Dog Name:')
        dog_name_label.grid(row=1, column=0)
        dog_name_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        dog_name_entry.grid(row=1, column=1)

        breed_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Breed:')
        breed_label.grid(row=2, column=0)
        breed_label_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        breed_label_entry.grid(row=2, column=1)

        colour_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Colour:')
        colour_label.grid(row=3, column=0)
        colour_label_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        colour_label_entry.grid(row=3, column=1)

        sex_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Sex:')
        sex_label.grid(row=4, column=0)
        sex_label_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        sex_label_entry.grid(row=4, column=1)

        year_of_birth_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Year of Birth:')
        year_of_birth_label.grid(row=5, column=0)
        year_of_birth_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        year_of_birth_entry.grid(row=5, column=1)


        number_of_dogs_label = Label(LeftFrame, font=('arial', 24, 'bold'), text='Number of Dogs')
        number_of_dogs_label.grid(row=6, column=0)
        number_of_dogs_entry = Entry(LeftFrame, font=('arial', 24, 'bold'))
        number_of_dogs_entry.grid(row=5, column=1)


    #Create the buttons

        add_button = Button(BottomFrame, pady = 1, bd = 4, font=('arial', 40, 'bold'), width=11, height=1, text='Add Data')
        add_button.grid(row=0, column=0,padx=3)

        update_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'), width=11, height=1, text='Update')
        update_button.grid(row=0, column=1,padx=3)

        plot_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'), width=11, height=1, text='Plot Graph')
        plot_button.grid(row=0, column=2,padx=3)

        reset_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'), width=11, height=1, text='Reset')
        reset_button.grid(row=0, column=3,padx=3)

        exit_button = Button(BottomFrame, pady=1, bd=4, font=('arial', 40, 'bold'), width=11, height=1, text='EXIT')
        exit_button.grid(row=0, column=4,padx=3)

    #Create the Treeview widget to display the data

    #create TTK Style instance
        style = ttk.Style()

    #incrase the fonts size of the treeview
        style.configure('Treeview.Heading', font=('TKDefaultFont', 18))
        style.configure('Treeview', rowheight = 40, font=('TKDefaultFont', 18))

        treeview_columns = ('Dog ID', 'Dogs Name', 'Breed', 'Colour', 'Sex', 'Year of birth', 'Number of Dogs')
        treeview = ttk.Treeview(RightFrame, columns = treeview_columns, show ='headings', height=10 )
        treeview.grid(row=0, columnspan=10, padx=34)

    # Set up the Treeview colums
        for col in treeview_columns:
            treeview.heading(col, text=col)
            treeview.column(col, width = 170)
            treeview.column(col, anchor ='center')

    #Load data from xlsx and display in the Treeview

        try:
            df = pd.read_excel('Rescue_Dogs.xlsx')
            for index, row in df.iterrows():
                    treeview.insert('','end', values=(
                    row['Dog_ID'],
                    row['Dog_Name'],
                    row['Breed'],
                    row['Colour'],
                    row['Sex'],
                    row['Year_of_Birth'],
                    row['Number_of_Dogs']))

        except Exception as e:
            messagebox.showerror('Error', str(e))

        #function to handle the Treeview row selection event
        def on_treeview_select(event):
            selected_item = treeview.focus()
            if selected_item:
                values = treeview.item(selected_item, 'values')
                dog_id_entry.delete(0, tk.END)
                dog_name_entry.delete(0, tk.END)
                breed_label_entry.delete(0, tk.END)
                colour_label_entry.delete(0, tk.END)
                sex_label_entry.delete(0, tk.END)
                year_of_birth_entry.delete(0, tk.END)
                number_of_dogs_entry.delete(0, tk.END)

                dog_id_entry.insert(0, values[0])
                dog_name_entry.insert(0, values[1])
                breed_label_entry.insert(0, values[2])
                colour_label_entry.insert(0, values[3])
                sex_label_entry.insert(0, values[4])
                year_of_birth_entry.insert(0, values[5])
                number_of_dogs_entry.insert(0, values[6])




        #Bind function of treeview selection event, amire kattintok kijelöli és bedobja az adatokat
        treeview.bind('<<TreeviewSelect>>', on_treeview_select)




if __name__ =='__main__':
    root = Tk()
    application = RescueDB (root)
    root.mainloop()




















