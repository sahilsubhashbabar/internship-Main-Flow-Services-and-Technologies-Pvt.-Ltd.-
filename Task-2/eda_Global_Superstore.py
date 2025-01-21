import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap

# Load dataset
file_path = r"C:\Users\sahil\Downloads\archive_d\newdata.xlsx"
df = pd.read_excel(file_path)

# Initial exploration
print("First few rows of the dataset:")
print(df.head())
print("\nDataset Info:")
print(df.info())
print("\nSummary Statistics:")
print(df.describe())

# Check for missing values
print("\nMissing Values per Column:")
print(df.isnull().sum())

# Handle missing values
# Filling numerical column 'sales' with mean
df['sales'] = df['sales'].fillna(df['sales'].mean())
# Filling categorical column 'region' with "Unknown"
df['region'] = df['region'].fillna("Unknown")

# Remove duplicates
before = len(df)
df = df.drop_duplicates()
after = len(df)
print(f"Removed {before - after} duplicate rows.")

# Handle outliers in the 'sales' column using IQR
Q1 = df['sales'].quantile(0.25)
Q3 = df['sales'].quantile(0.75)
IQR = Q3 - Q1
lower_bound = Q1 - 1.5 * IQR
upper_bound = Q3 + 1.5 * IQR
df = df[(df['sales'] >= lower_bound) & (df['sales'] <= upper_bound)]
print("Outliers removed based on IQR.")

# Statistical summary for relevant numerical columns
print("\nDescriptive Statistics:")
print(df[['sales', 'expenses', 'revenue ']].describe())

# Compute correlation matrix for numerical columns
numeric_df = df.select_dtypes(include=[np.number])
correlation = numeric_df.corr()
print("\nCorrelation Matrix:")
print(correlation)

# Visualizations
# 1. Histogram for Sales Distribution
plt.figure(figsize=(10, 6))
df['sales'].hist(bins=30, color='peachpuff', edgecolor='black')
plt.title('Sales Distribution', fontsize=16, weight='bold')
plt.xlabel('Sales', fontsize=12)
plt.ylabel('Frequency', fontsize=12)
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()

# 2. Boxplot for Sales
plt.figure(figsize=(10, 6))
sns.boxplot(x=df['sales'], color='lightseagreen')
plt.title('Boxplot of Sales', fontsize=16, weight='bold')
plt.xlabel('Sales', fontsize=12)
plt.grid(True, linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()

# 3. Heatmap for Correlation Matrix
colors = ["#ffb6c1", "#8a2be2"]
cmap = LinearSegmentedColormap.from_list("pink_violet", colors)
plt.figure(figsize=(8, 8))
sns.heatmap(correlation, annot=True, cmap=cmap, fmt=".2f", annot_kws={"size": 12, 'weight': 'bold'}, cbar_kws={'shrink': 0.8})
plt.title('Correlation Heatmap', fontsize=16, weight='bold')
plt.tight_layout()
plt.show()

# 4. Barplot for Total Sales by Region
plt.figure(figsize=(10, 6))
sns.barplot(x='region', y='sales', data=df, hue='region', palette='Set2')
plt.title('Total Sales by Region', fontsize=16, weight='bold')
plt.xlabel('Region', fontsize=12)
plt.ylabel('Total Sales', fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()
