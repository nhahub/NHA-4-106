import streamlit as st
import pandas as pd
import joblib

# تحميل الموديل والـ Scaler
model = joblib.load('kmeans_model.pkl')
scaler = joblib.load('scaler.pkl')

st.title("🚛 Trip Clustering Predictor")
st.write("Enter the trip details to find out which category (Cluster) it belongs to")

dist = st.number_input("(Miles)", min_value=0.0, value=500.0)
dur = st.number_input("(Hours)", min_value=0.0, value=10.0)
weight = st.number_input("(Lbs)", min_value=0.0, value=20000.0)
idle = st.number_input("(Idle Hours)", min_value=0.0, value=2.0)

if st.button("Category expectation"):
    data = pd.DataFrame([[dist, dur, weight, idle]], 
                        columns=['actual_distance_miles', 'actual_duration_hours', 'weight_lbs', 'idle_time_hours'])
    
    scaled_data = scaler.transform(data)
    prediction = model.predict(scaled_data)[0]
    
    
    categories = {
        0: "Light transport trips (medium distances)",
        1: "Very long distance flights",
        2: "Heavy transport trips"
    }
    st.success(f"This trip is affiliated with: {categories[prediction]}"))