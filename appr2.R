library(readxl)
library(dplyr)
library(lubridate)
library(writexl)

# Read and process markers data
markers_file_path <- "/Users/williambell/Downloads/Marker_Data.xlsx"
# Read the data without column names
markers_data_frame <- read_excel(markers_file_path, col_names = FALSE)
# Replace row 1 from column 4 onwards with values from row 2
markers_data_frame[1, 4:ncol(markers_data_frame)] <- markers_data_frame[2, 4:ncol(markers_data_frame)]
# Remove row 2
markers_data_frame <- markers_data_frame[-2, ]
# Convert all columns from 4 onward to character
markers_data_frame <- markers_data_frame %>% mutate(across(4:ncol(.), as.character))

# Function to convert Excel numeric dates to datetime format matching schedule_data_frame
excel_num_to_datetime <- function(x) {
  # If the value is NA or empty, return the original value
  if (is.na(x) || x == "" || is.null(x)) {
    return(as.character(x))
  }
  # Try to handle different input types
  tryCatch({
    # If it's already in date format (like "2025-07-07"), parse it directly
    if (grepl("-", x)) {
      datetime <- as.POSIXct(x)
    } else {
      # Try to convert to numeric - if it fails, return original value
      excel_num <- suppressWarnings(as.numeric(x))
      if (is.na(excel_num)) {
        return(as.character(x))  # Return original value if not numeric
      }
      # Convert Excel numeric date to R datetime
      datetime <- as.POSIXct((excel_num - 25569) * 86400, origin = "1970-01-01", tz = "UTC")
    }
    # Format as "YYYY-MM-DD 00:00:00"
    format(datetime, "%Y-%m-%d %H:%M:%S")
  }, error = function(e) {
    # If any error occurs, return the original value
    return(as.character(x))
  })
}

# Apply the conversion to all columns from 4 onwards
markers_data_frame <- markers_data_frame %>% 
  mutate(across(4:ncol(.), ~ sapply(., excel_num_to_datetime)))

# Set the first row as column headers
colnames(markers_data_frame) <- as.character(markers_data_frame[1, ])
# Remove the first row (now that it's the header)
markers_data_frame <- markers_data_frame[-1, ]

# Read schedule data
schedule_file_path <- "/Users/williambell/Downloads/School_Test_Delivery_Schedule.xlsx"
# Read the data with column names
schedule_data_frame <- read_excel(schedule_file_path, col_names = TRUE)

# Function to get available dates for a marker (dates where they have "Y")
get_marker_available_dates <- function(marker_name, markers_df) {
  marker_row <- markers_df[markers_df$`Marker Name` == marker_name, ]
  if (nrow(marker_row) == 0) return(character(0))
  # Get date columns (columns 4 onwards)
  date_cols <- colnames(markers_df)[4:ncol(markers_df)]
  available_dates <- character(0)
  for (date_col in date_cols) {
    if (!is.na(marker_row[[date_col]]) && marker_row[[date_col]] == "Y") {
      available_dates <- c(available_dates, date_col)
    }
  }
  return(available_dates)
}

# Function to find next available date for a marker after a given date
find_next_available_date <- function(marker_name, after_date, markers_df) {
  available_dates <- get_marker_available_dates(marker_name, markers_df)
  if (length(available_dates) == 0) return(NA)
  # Convert to dates for comparison
  available_dates_parsed <- as.Date(available_dates)
  after_date_parsed <- as.Date(after_date)
  # Find dates after the given date
  future_dates <- available_dates_parsed[available_dates_parsed > after_date_parsed]
  if (length(future_dates) == 0) return(NA)
  # Return the earliest future date
  return(min(future_dates))
}

# Function to calculate days between dates
days_between <- function(date1, date2) {
  as.numeric(as.Date(date2) - as.Date(date1))
}
# Initialize a global usage tracker (marker pair -> count)
global_usage_counts <- new.env(hash = TRUE)

# Function to update usage counts for a pair across all tables
update_global_usage <- function(pair_name) {
  # Initialize count if not exists
  if (!exists(pair_name, envir = global_usage_counts)) {
    assign(pair_name, 0, envir = global_usage_counts)
  }
  # Increment count
  assign(pair_name, get(pair_name, envir = global_usage_counts) + 1, envir = global_usage_counts)
}

# Function to get current usage count for a pair
get_usage_count <- function(pair_name) {
  if (exists(pair_name, envir = global_usage_counts)) {
    return(get(pair_name, envir = global_usage_counts))
  } else {
    return(0)
  }
}

# Modified create_marker_pair_table to use global usage counts
create_marker_pair_table <- function(school_serial, expected_delivery_date, markers_df) {
  marker_names <- markers_df$`Marker Name`
  n_markers <- length(marker_names)
  
  # Generate all unique pairs
  pairs <- list()
  pair_names <- character(0)
  for (i in 1:(n_markers-1)) {
    for (j in (i+1):n_markers) {
      marker_a <- marker_names[i]
      marker_b <- marker_names[j]
      # Check if both markers are available on any date after expected delivery date
      next_date_a <- find_next_available_date(marker_a, expected_delivery_date, markers_df)
      next_date_b <- find_next_available_date(marker_b, expected_delivery_date, markers_df)
      if (!is.na(next_date_a) && !is.na(next_date_b)) {
        pairs[[length(pairs) + 1]] <- c(marker_a, marker_b)
        pair_name <- paste0("(", marker_a, ", ", marker_b, ")")
        pair_names <- c(pair_names, pair_name)
      }
    }
  }
  
  if (length(pairs) == 0) {
    return(data.frame(
      `Pairs Available to Mark` = character(0),
      `Number of Days Until Markers Available` = character(0),
      `Sum of Days to Wait` = numeric(0),
      `Residual Sum` = numeric(0),
      `Usage Count` = numeric(0),
      `Suitability Score` = numeric(0),
      stringsAsFactors = FALSE
    ))
  }
  
  # Calculate days until available for each pair
  days_until_available <- character(0)
  sum_days <- numeric(0)
  for (i in 1:length(pairs)) {
    marker_a <- pairs[[i]][1]
    marker_b <- pairs[[i]][2]
    next_date_a <- find_next_available_date(marker_a, expected_delivery_date, markers_df)
    next_date_b <- find_next_available_date(marker_b, expected_delivery_date, markers_df)
    days_a <- days_between(expected_delivery_date, next_date_a)
    days_b <- days_between(expected_delivery_date, next_date_b)
    days_until_available <- c(days_until_available, paste0("(", days_a, ", ", days_b, ")"))
    sum_days <- c(sum_days, days_a + days_b)
  }
  
  # Calculate residual sums (max sum - current sum)
  max_sum <- ifelse(length(sum_days) > 0, max(sum_days), 0)
  residual_sums <- max_sum - sum_days
  
  # Get usage counts from global tracker
  usage_counts <- sapply(pair_names, get_usage_count)
  
  # Calculate Suitability Score (Residual Sum - N * Usage Count) --- Choose N as you like depending on how heavily you want to penalise repetition.
  suitability_scores <- residual_sums - (10 * usage_counts)
  
  # Create the table
  result_table <- data.frame(
    `Pairs Available to Mark` = pair_names,
    `Number of Days Until Markers Available` = days_until_available,
    `Sum of Days to Wait` = sum_days,
    `Residual Sum` = residual_sums,
    `Usage Count` = usage_counts,
    `Suitability Score` = suitability_scores,
    stringsAsFactors = FALSE
  )
  
  return(result_table)
}

# Create tables for each School Serial No. and track best pairs
school_tables <- list()
best_pairs <- data.frame(
  `School Serial No.` = character(),
  `Best Marker Pair` = character(),
  `Suitability Score` = numeric(),
  `Usage Count` = numeric(),
  stringsAsFactors = FALSE
)

for (i in 1:nrow(schedule_data_frame)) {
  school_serial <- schedule_data_frame$`School Serial No.`[i]
  expected_delivery_date <- as.Date(schedule_data_frame$`Expected Test Delivery Date`[i])
  
  # Create table for this school (with current usage counts)
  school_table <- create_marker_pair_table(school_serial, expected_delivery_date, markers_data_frame)
  
  # Store in list with school serial as name
  school_tables[[school_serial]] <- school_table
  
  # Find the best pair(s) if any available
  if (nrow(school_table) > 0) {
    max_score <- max(school_table$Suitability.Score)
    best_candidates <- school_table[school_table$Suitability.Score == max_score, ]
    
    # Randomly select one if there are ties
    selected_pair <- if (nrow(best_candidates) > 1) {
      set.seed(123)
      best_candidates[sample(nrow(best_candidates), 1), ]
    } else {
      best_candidates
    }
    
    # Update global usage count for this pair
    update_global_usage(selected_pair$Pairs.Available.to.Mark)
    
    # Add to best_pairs data frame (with updated usage count)
    best_pairs <- rbind(best_pairs, data.frame(
      `School Serial No.` = school_serial,
      `Best Marker Pair` = selected_pair$Pairs.Available.to.Mark,
      `Suitability Score` = selected_pair$Suitability.Score,
      `Usage Count` = get_usage_count(selected_pair$Pairs.Available.to.Mark),
      stringsAsFactors = FALSE
    ))
  }
  
  # Also create individual variable for easy access
  assign(paste0("table_", school_serial), school_table, envir = .GlobalEnv)
}


# Print all best pairs with updated usage counts
cat("\nBest Marker Pairs for Each School (with usage tracking):\n")
print(best_pairs)

# Prepare the final output data frame with clean column names
final_output <- data.frame(
  `School Serial No.` = character(),
  `Expected Test Delivery Date` = character(),  # Store as character to control formatting
  `Marker 1` = character(),
  `Marker 2` = character(),
  `Marker 3` = character(),
  `Marker 4` = character(),
  stringsAsFactors = FALSE,
  check.names = FALSE
)

# Populate the final output data frame
for (i in 1:nrow(best_pairs)) {
  school_serial <- best_pairs$School.Serial.No.[i]
  pair <- best_pairs$Best.Marker.Pair[i]
  
  # Extract marker names from the pair string "(marker1, marker2)"
  markers <- gsub("[()]", "", pair) %>% 
    strsplit(", ") %>% 
    unlist()
  
  # Get and format the expected delivery date as DD/MM/YYYY
  delivery_date <- schedule_data_frame$`Expected Test Delivery Date`[
    schedule_data_frame$`School Serial No.` == school_serial
  ]
  formatted_date <- format(as.Date(delivery_date), "%d/%m/%Y")
  
  # Add to final output
  final_output <- rbind(final_output, data.frame(
    `School Serial No.` = school_serial,
    `Expected Test Delivery Date` = formatted_date,
    `Marker 1` = ifelse(length(markers) >= 1, markers[1], NA),
    `Marker 2` = ifelse(length(markers) >= 2, markers[2], NA),
    `Marker 3` = NA,
    `Marker 4` = NA,
    stringsAsFactors = FALSE,
    check.names = FALSE
  ))
}

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet
addWorksheet(wb, "Marker Assignments")

# Write the data to the worksheet
writeData(wb, sheet = 1, final_output, startCol = 1, startRow = 1, colNames = TRUE, rowNames = FALSE)

# Set column widths
setColWidths(wb, sheet = 1, cols = 1:6, widths = c(15, 23, 15, 15, 15, 15))

# Create a bold style for headers
header_style <- createStyle(textDecoration = "bold")

# Apply bold style to the header row
addStyle(wb, sheet = 1, style = header_style, rows = 1, cols = 1:6)

# Freeze the header row
freezePane(wb, sheet = 1, firstRow = TRUE)

# Save the workbook with proper column widths
output_file_path <- "/Users/williambell/Downloads/Marker_Assignments.xlsx"
saveWorkbook(wb, output_file_path, overwrite = TRUE)

cat("\nFinal output saved to:", output_file_path, "\n")