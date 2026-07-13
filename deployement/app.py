import streamlit as st
import pandas as pd
import joblib

st.set_page_config(page_title="Logistics AI Dashboard", layout="wide")

st.title("🚛 Logistics Decision Intelligence")
st.write("Professional AI-powered tools for supply chain operations.")


kmeans_model = joblib.load('kmeans.pkl')
scaler = joblib.load('scaler.pkl')
driver_model = joblib.load('Efficiency.pkl')

tab1, tab2 = st.tabs(["Trip Segmentation (Clustering)", "Driver Efficiency Prediction"])

with tab1:
    st.header("Trip Category Predictor")
    col1, col2 = st.columns(2)
    
    with col1:
        dist = st.number_input("Distance (Miles)", min_value=0.0, value=500.0, key="dist_c")
        dur = st.number_input("Duration (Hours)", min_value=0.0, value=10.0, key="dur_c")
    with col2:
        weight = st.number_input("Weight (Lbs)", min_value=0.0, value=20000.0, key="weight_c")
        idle = st.number_input("Idle Time (Hours)", min_value=0.0, value=2.0, key="idle_c")
        
    if st.button("Predict Category"):
        data = pd.DataFrame([[dist, dur, weight, idle]], 
                            columns=['actual_distance_miles', 'actual_duration_hours', 'weight_lbs', 'idle_time_hours'])
        scaled_data = scaler.transform(data)
        prediction = kmeans_model.predict(scaled_data)[0]
        
        categories = {
            0: "Light transport trips (medium distances)",
            1: "Very long distance trips",
            2: "Heavy transport trips"
        }
        st.success(f"Result: {categories[prediction]}")

with tab2:
    st.header("Driver Efficiency Predictor")
    st.write("Predict estimated MPG based on input metrics.")
    
    col1, col2 = st.columns(2)
    with col1:
        dist_r = st.number_input("Total Monthly Miles", value=5000, key="dist_r")
        dur_r = st.number_input("Total Monthly Hours", value=150, key="dur_r")
    with col2:
        weight_r = st.number_input("Average Load Weight", value=20000, key="weight_r")
        trips_r = st.number_input("Total Trips Count", value=20, key="trips_r")
        
    if st.button("Predict Efficiency"):
       
        data_r = pd.DataFrame([[dist_r, dur_r, weight_r, trips_r]], 
                            columns=['actual_distance_miles', 'actual_duration_hours', 'weight_lbs', 'trip_id'])
        
        
        predicted_fuel = driver_model.predict(data_r)[0]
        
        
        if predicted_fuel > 0:
            mpg = dist_r / predicted_fuel
            st.metric("Predicted Efficiency (MPG)", f"{mpg:.2f}")
        else:
            st.error("Could not calculate efficiency (fuel prediction was invalid).")
