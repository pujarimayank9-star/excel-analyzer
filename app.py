# -------- DEEP DATASET OBSERVATION + ASPECT RECOMMENDATIONS --------

st.header("Deep Dataset Observations")

deep_observations = []

# Numeric variability observation
if len(numeric_cols) > 0:
    variability = df[numeric_cols].std().mean()
    avg_value = df[numeric_cols].mean().mean()

    if variability > avg_value * 0.3:
        deep_observations.append(
            "Across numeric metrics the dataset shows noticeable variability, "
            "indicating that performance or values fluctuate significantly "
            "between observations."
        )
    else:
        deep_observations.append(
            "Numeric values appear relatively stable, suggesting consistent "
            "performance or measurement across the dataset."
        )

# Correlation observation
if len(numeric_cols) > 1:
    corr = df[numeric_cols].corr().abs()
    corr_values = corr.values.flatten()
    corr_values = corr_values[corr_values < 1]

    if len(corr_values) > 0:
        strongest = corr_values.max()

        if strongest > 0.7:
            deep_observations.append(
                "Some metrics show strong relationships with each other, "
                "meaning changes in one variable may significantly influence another."
            )
        elif strongest > 0.4:
            deep_observations.append(
                "Several variables demonstrate moderate interaction, suggesting "
                "partial dependency between some aspects of the dataset."
            )
        else:
            deep_observations.append(
                "Most variables appear largely independent, meaning performance "
                "or outcomes in one metric do not strongly influence others."
            )

# Category distribution observation
for cat in cat_cols:
    counts = df[cat].value_counts()
    if len(counts) > 0:
        top_share = counts.max() / counts.sum()

        if top_share > 0.5:
            deep_observations.append(
                f"The distribution of {cat} is heavily concentrated in one category, "
                f"indicating imbalance in the dataset."
            )
        elif top_share > 0.3:
            deep_observations.append(
                f"{cat} shows moderate concentration around a few categories."
            )
        else:
            deep_observations.append(
                f"{cat} appears fairly balanced across multiple categories."
            )

for obs in deep_observations:
    st.write("Observation:", obs)


# -------- ASPECT BASED STRATEGIC RECOMMENDATIONS --------

st.header("Strategic Recommendations")

aspect_recommendations = []

# Recommendation based on variability
if len(numeric_cols) > 0:
    variability = df[numeric_cols].std().mean()
    avg_value = df[numeric_cols].mean().mean()

    if variability > avg_value * 0.3:
        aspect_recommendations.append(
            "Standardize processes or evaluation criteria to reduce variability "
            "and improve consistency across metrics."
        )
    else:
        aspect_recommendations.append(
            "Maintain current operational or evaluation structure since "
            "performance appears consistent across the dataset."
        )

# Recommendation based on correlations
if len(numeric_cols) > 1:
    corr = df[numeric_cols].corr().abs()
    corr_values = corr.values.flatten()
    corr_values = corr_values[corr_values < 1]

    if len(corr_values) > 0:
        strongest = corr_values.max()

        if strongest > 0.6:
            aspect_recommendations.append(
                "Leverage the relationships between strongly correlated metrics "
                "to design integrated improvement strategies."
            )
        else:
            aspect_recommendations.append(
                "Since most metrics are weakly related, improvements may need "
                "to be targeted independently for each aspect."
            )

# Recommendation based on categorical imbalance
for cat in cat_cols:
    counts = df[cat].value_counts()
    if len(counts) > 0:
        top_share = counts.max() / counts.sum()

        if top_share > 0.5:
            aspect_recommendations.append(
                f"Consider balancing representation within {cat} to avoid "
                f"over-reliance on a single dominant category."
            )
        else:
            aspect_recommendations.append(
                f"The distribution of {cat} appears balanced, so strategies "
                f"can focus on improving quality rather than structural changes."
            )

for rec in aspect_recommendations:
    st.write("Recommendation:", rec)
