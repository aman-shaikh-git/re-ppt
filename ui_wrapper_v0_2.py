import streamlit as st
import os
from pathlib import Path
import time 
import tkinter as tk
from tkinter import filedialog
import read_scorecard_v1_1 as mapper
import generate_scorecards_v1_1 as generator

def get_save_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    root.attributes("-topmost", True)  # Bring the dialog to the front
    file_path = filedialog.asksaveasfilename(
        defaultextension=".pptx",
        filetypes=[("PowerPoint files", "*.pptx")],
        title="Select Location"
    )
    root.destroy()
    return file_path

st.set_page_config(page_title="Re-PPT App_v1.0", layout="centered")

st.title("Re-PPT")
st.markdown("---")

# Use tabs to separate the two-step workflow
tab1, tab2 = st.tabs(["1. Create Mapping Excel", "2. Generate Scorecards"])

with tab1:
    st.header("Step 1: Map Template")
    st.info("Upload your template slide to generate the Excel input sheet.")
    
    uploaded_pptx = st.file_uploader("Upload Template PPTX", type=["pptx"], key="mapper_pptx")
    
    if uploaded_pptx:
        # Save temp file
        temp_path = Path("temp_template.pptx")
        with open(temp_path, "wb") as f:
            f.write(uploaded_pptx.getbuffer())
            
        if st.button("Generate Excel Map"):
            with st.spinner("PowerPoint is mapping shapes..."):
                # Call your existing mapper logic
                # Modified slightly to return the filename
                excel_out = mapper.process_ppt_template(temp_path)
                
            if os.path.exists(excel_out):
                with open(excel_out, "rb") as f:
                    st.download_button(
                        label="Download Excel Input Sheet",
                        data=f,
                        file_name=excel_out,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                st.success("Mapping Complete!")

with tab2:
    st.header("Step 2: Generate Deck")
    st.info("Upload your original template and your edited Excel sheet.")
    
    gen_pptx = st.file_uploader("Upload Template PPTX", type=["pptx"], key="gen_pptx")
    gen_xlsx = st.file_uploader("Upload Edited Excel", type=["xlsx"], key="gen_xlsx")
    
    if gen_pptx and gen_xlsx:
        if st.button("Create Final Scorecards"):
            #Ask user for location to save outputs in
            output_path_str = get_save_path()

            if not output_path_str:
                st.warning("Action Cancelled: No save location selected")
                st.stop() # End script if no location selected
        
            #Start timer
            start_time = time.time()
            # Save temp files
            t_path = Path("temp_gen_template.pptx")
            e_path = Path("temp_gen_data.xlsx")

            #Ask user for location to save
            
            with open(t_path, "wb") as f: f.write(gen_pptx.getbuffer())
            with open(e_path, "wb") as f: f.write(gen_xlsx.getbuffer())
            
            with st.spinner("Duplicating slides and filling data..."):
                final_pptx, slide_count = generator.generate_deck(t_path, e_path, output_path_str)

            #End timer
            runtime_duration = round(time.time() - start_time, 2)

            if os.path.exists(final_pptx):
                with open(final_pptx, "rb") as f:
                    st.download_button(
                        label="Download Final Presentation",
                        data=f,
                        file_name="Re-PPT_Generated_Output.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                st.success(f"Generation Complete! {slide_count} slides generated in {runtime_duration} seconds.")
                st.info(f"Saved to: {output_path_str}")


    st.sidebar.markdown("---")
    st.sidebar.subheader("Application Controls")

    if st.sidebar.button("Exit and Close RePPT"):
        st.sidebar.warning("Shutting down...You can now safely close this tab. ")
        time.sleep(2)
        # This kills the specific process running the server
        os._exit(0)
        os.kill(os.getpid(), signal.SIGTERM)