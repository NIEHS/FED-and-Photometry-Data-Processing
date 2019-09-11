# FED and Photometry Data Processing

# install/load required packages
for (pkg in c("ggplot2", "ggrepel", "gridExtra", "data.table", "readxl", "writexl", "R.utils",
              "shiny", "shinycssloaders", "DT", "rhandsontable", "shinyBS")) {
    if (!require(pkg, character.only = T)) {
        install.packages(pkg)
        library(pkg, character.only = T)
    }
}

# some options and default values
show.debug.messages = T
options(shiny.maxRequestSize=100*1024^2)

default = list(
    input_sheet = 1,
    
    event_column = "FED",
    data_columns = c("Gcamp/Tdtomato"),
    
    frequency = 25,
    event_offset = -0.5,
    periods = data.frame(
        name = c("unassigned", "baseline", "approach", "consumption", "post consumption"),
        start = c(-45, -45, -2, 5, 30),
        end = c(45, -30, 0, 10, 45)),
    
    zscore_reference = c("(entire event)"),
    pdff_reference = c("baseline"),
    pdff_statistic = c("mean"),
    
    plot_overview_min = 0,
    plot_overview_max = 0,
    plot_overview_points = T,
    plot_overview_lines = T,
    
    plot_density_bins = 180,
    
    plot_events_points = T,
    plot_events_lines = T,
    plot_events_shaded = T,
    
    plot_summary_type = "Raw",
    plot_summary_points = T,
    plot_summary_labels = T,
    
    plot_gradient = c(
        "#f1eef6", 
        "#bdc9e1", 
        "#74a9cf", 
        "#2b8cbe", 
        "#045a8d"
    )
)

# ---------------------------------------------------------------------------------------------------------------------

dbg = function(...) if (show.debug.messages) printf(...)

MinMeanSEMMax <- function(x) {
    v <- c(min(x), mean(x)-sd(x)/sqrt(length(x)), mean(x), mean(x)+sd(x)/sqrt(length(x)), max(x))
    names(v) <- c("ymin", "lower", "middle", "upper", "ymax")
    v
}

# Shiny UI Definitions
ui = fluidPage(
    titlePanel("FED and Photometry Data Processing"),
    sidebarLayout(
        sidebarPanel(
            fileInput("input_file", ".XLSX Input File", accept = c("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx")),
            tags$hr(),
            selectInput("event_column", "Event Column", choices = c()),
            checkboxGroupInput("data_columns", "Data Column(s)", choices = c()),
            actionButton("process_button", label = "Update"),
            tags$hr(),
            downloadButton("output_download", label = "Export Processed .XLSX File")
        ),
        mainPanel(
            tabsetPanel(
                tabPanel(
                    "Configuration",
                    tags$br(),
                    fluidRow(
                        column(
                            tipify(
                                numericInput("input_sheet", "Sheet Number", default$input_sheet),
                                title = "The sheet number (starting from 1) that contains the data to read within the selected Excel file.",
                                placement = "left"),
                            width = 4
                        ),
                        column(
                            tipify(
                                numericInput("event_offset", "Event Offset (s)", default$event_offset, step = 0.001),
                                title = "Amount of time in seconds to adjust each event by in order to account for delays between the actual event and corresponding FED signal.",
                                placement = "left"),
                            width = 4
                        ),
                        column(
                            tipify(
                                numericInput("frequency", "Data Frequency (Hz)", default$frequency, min = 1), 
                                title = "Number of observations recorded per second.",
                                placement = "left"),
                                width = 4
                            )
                    ),
                    tags$hr(),
                    fluidRow(
                        column(
                            tags$b("Event Interval Assignments"),
                            withSpinner(rHandsontableOutput("period_table")),
                            bsTooltip("period_table", "List of named intervals to assign surrounding each event. Assignments are done in order of appearance, so any intervals that are overlapped by another will be overwritten. The \"unassigned\" interval should come first, as its span will determine the length of each event. It is also excluded from summary value tables.", "left"),
                            width = 6
                        ),
                        column(
                            fileInput("period_file", "Load from .CSV", accept = c(".csv")),
                            downloadButton("period_download", "Save as .CSV"),
                            actionButton("period_reset", "Reset to Defaults"),
                            width = 6
                        )
                    ),
                    tags$hr(),
                    fluidRow(
                        column(
                            tipify(
                                selectInput("zscore_reference", "z-Score Reference Interval", default$zscore_reference, selected = default$zscore_reference),
                                title = "For each event, the mean and standard deviation of this interval are used when calculating z-scores.",
                                placement = "left"),
                            width = 6
                        )
                    ),
                    fluidRow(
                        column(
                            tipify(
                                selectInput("pdff_reference", "%dF/F Reference Interval", default$pdff_reference, selected = default$pdff_reference),
                                title = "For each event, the interval to be used as the reference when calculating %dF/F values.",
                                placement = "left"),
                            width = 6
                        ),
                        column(
                            tipify(
                                selectInput("pdff_statistic", "%dF/F Reference Statistic", c("mean", "median"), selected = default$pdff_statistic),
                                title = "The statistic to be taken over the reference interval when calculating %dF/F values.",
                                placement = "left"),
                            width = 6
                        )
                    )
                ),
                tabPanel(
                    "Data",
                    tags$br(),
                    fluidRow(
                        column(
                            numericInput("overview_min", "Minimum", default$plot_overview_min),
                            width = 3
                        ),
                        column(
                            numericInput("overview_max", "Maximum", default$plot_overview_max),
                            actionButton("overview_reset", "Reset"),
                            width = 3
                        ),
                        column(
                            checkboxGroupInput("overview_options", "Options", choices = c("Points", "Lines", "Event Markers"), selected = c("Points", "Lines", "Event Markers"), inline = F),
                            width = 6
                        )
                    ),
                    tags$hr(),
                    div(withSpinner(plotOutput("overview_plot", width = "80%", height = "auto")), align = "center")
                ),
                tabPanel("Density",
                         div(withSpinner(plotOutput("density_plot", width = "80%", height = "auto")), align = "center")
                ),
                tabPanel("Events",
                        tags$br(),
                        fluidRow(
                            column(
                                checkboxGroupInput("events_options", "Options", choices = c("Points", "Lines", "Shade Intervals"), selected = c("Points", "Lines", "Shade Intervals"), inline = F),
                                width = 4
                            ),
                            column(DTOutput("events_table"), width = 8)
                        ),
                        tags$hr(),
                        div(withSpinner(plotOutput("events_plot", width = "80%")), align = "center")
                ),
                tabPanel("Summary",
                        fluidRow(
                            column(selectInput("summary_type", "Summary", c("Raw", "%dF/F", "z Score")), width = 3),
                            column(checkboxGroupInput("summary_options", "Options", inline = T,
                                                      choices = c("Points", "Labels"),
                                                      selected = c("Points", "Labels")), width = 9)
                        ),
                        div(withSpinner(plotOutput("summary_boxplot", width = "80%", height = "400px")), align = "center"),
                        withSpinner(rHandsontableOutput("summary_table"))
                )
            ),
        style = "z-index:500;") # bstooltip issues otherwise?
    )
)

# ---------------------------------------------------------------------------------------------------------------------

server = function(input, output, session) {
    data = reactiveValues()
    config = reactiveValues()
    
    # setup configuration from default values
    for (v in names(default))
        config[[v]] = default[[v]]
    
    # process events for changing file options
    observeEvent(input$input_sheet, {
        config$input_sheet = input$input_sheet
        dbg("set input file sheet number to %d\n", input$input_sheet)
    })
    
    observeEvent(input$event_offset, {
        config$event_offset = input$event_offset
        dbg("set event offset to %f\n", input$event_offset)
    })
    
    observeEvent(input$frequency, {
        config$frequency = input$frequency
        dbg("set frequency to %f\n", input$frequency)
    })
    
    
    # process events for changes to interval definitions
    observeEvent(input$period_table, {
        config$periods = hot_to_r(input$period_table)
        intervals = as.character(config$periods[, "name"])
        intervals = intervals[intervals != "unassigned"]
        updateSelectInput(session, "zscore_reference", choices = cbind("(entire event)", intervals))
        updateSelectInput(session, "pdff_reference", choices = intervals)
        dbg("updated interval configuration\n")
    })
    
    observeEvent(input$period_file, {
        config$periods = read.csv(input$period_file$datapath, stringsAsFactors = F)
        intervals = as.character(config$periods[, "name"])
        intervals = intervals[intervals != "unassigned"]
        updateSelectInput(session, "zscore_reference", choices = cbind("(entire event)", intervals))
        updateSelectInput(session, "pdff_reference", choices = intervals)
        dbg("updated intervals from uploaded file\n")
    })
    
    observeEvent(input$period_reset, {
        config$periods = default$periods
        intervals = as.character(config$periods[, "name"])
        intervals = intervals[intervals != "unassigned"]
        updateSelectInput(session, "zscore_reference", choices = cbind("(entire event)", intervals))
        updateSelectInput(session, "pdff_reference", choices = intervals)
        dbg("reset intervals to original values\n")
    })
    
    # statistic reference interval options
    observeEvent(input$zscore_reference, {
        config$zscore_reference = input$zscore_reference
        dbg("updated z-score reference interval: %s\n", input$zscore_reference)
    })
    
    observeEvent(input$pdff_reference, {
        config$pdff_reference = input$pdff_reference
        dbg("updated %%dF/F reference interval: %s\n", input$pdff_reference)
    })
    
    observeEvent(input$pdff_statistic, {
        config$pdff_statistic = input$pdff_statistic
        dbg("updated %%dF/F reference statistic: %s\n", input$pdff_statistic)
    })
    
    observeEvent(input$overview_min, {
        config$overview_min = input$overview_min
        dbg("set overview plot minimum to %f\n", config$overview_min)
    })
    
    observeEvent(input$overview_max, {
        config$overview_max = input$overview_max
        dbg("set overview plot maximum to %f\n", config$overview_max)
    })
    
    observeEvent(input$overview_reset, {
        updateNumericInput(session, "overview_min", value = 0)
        updateNumericInput(session, "overview_max", value = data$input[, .N] / config$frequency)
        dbg("reset overview limits\n")
    })
    
    # event variable selection
    observeEvent(input$event_column, {
        req(data$input)
        config$event_column = input$event_column
        dbg("set event column to %s\n", input$event_column)
    })
    
    # data variable(s) selection
    observeEvent(input$data_columns, {
        req(data$input)
        config$data_columns = input$data_columns
        dbg("set data columns to %s\n", paste(input$data_columns, collapse = ", "))
    })
    
    # read xlsx to data.table on upload
    observeEvent(c(input$input_file, input$frequency), {
        req(input$input_file)
        dt = data.table(read_xlsx(input$input_file$datapath, sheet = config$input_sheet))
        dt[, time := 1:dt[, .N] / config$frequency]
        data$input = dt
        dbg("file uploaded has %d rows, %d columns, spans %f seconds at %d Hz\n", dim(dt)[1], dim(dt)[2], dt[, .N] / config$frequency, config$frequency)
        updateNumericInput(session, "overview_min", value = 0)
        updateNumericInput(session, "overview_max", value = dt[, .N] / config$frequency)
    })
    
    # populate event and data column names from uploaded file
    observeEvent(data$input, {
        updateSelectInput(session, "event_column", choices = names(data$input))
        updateCheckboxGroupInput(session, "data_columns", choices = names(data$input))
        
        if (default$event_column %in% names(data$input)) {
            updateSelectInput(session, "event_column", selected = grep(default$event_column, names(data$input), value = T))
        }
        
        if  (all(default$data_columns %in% names(data$input))) {
            updateSelectInput(session, "data_columns", selected = grep(default$data_columns, names(data$input), value = T))
        }
        
        dbg("updated column choices (%s)\n", paste(names(data$input), collapse = ", "))
    })
    
    # identify TTL-type events as having occured where the variable changes from a 0 to 1
    # also adjust all event times by event_offset option
    observeEvent(c(data$input, input$event_column, input$event_offset), {
        req(data$input, input$event_column)
        if (input$event_column %in% names(data$input)) {
            events = data$input[shift(get(config$event_column), n=1) == 0 & get(config$event_column) == 1,]
            events[, time := time + config$event_offset]
            events[, event := .I]
            data$events = events
            dbg("identified %d events (%s)\n", events[, .N], paste(events[, time], collapse = ", "))
        }
    })
    
    # once dataset is uploaded and events are identified, process the data into another dataset containing each event's entire data series and relevant statistics
    observeEvent(c(data$events, input$process_button, input$zscore_reference, input$pdff_reference, input$pdff_statistic), {
        req(data$events)
        min.offset = min(unlist(config$periods))
        max.offset = max(unlist(config$periods))
        
        # revisit this?
        epsilon = (2/3) / config$frequency
        data$eventdata = data.table()
        
        for (i in 1:data$events[, .N]) {
            event = data$input[((time - data$events[i, time]) - min.offset >= -epsilon) & ((time - data$events[i, time]) - max.offset <= epsilon),]
            event[, `:=`(period = "", event = i, event_time = time - data$events[i, time])]
            
            for (j in 1:dim(config$periods)[1]) {
                event[(event_time - config$periods[j, "start"] >= - epsilon) & (event_time - config$periods[j, "end"] <= epsilon), `:=`(period = config$periods[j, "name"])]
            }
            
            for (var in input$data_columns) {
                new.vars = paste0(var, c("_mean_baseline", "_pdff", "_zscore", "_mean"))
                baseline = event[period == "baseline", list(mean = mean(get(var)), median = median(get(var)))]
                
                if (config$zscore_reference != "(entire event)") {
                    zscore_reference = event[period == config$zscore_reference, list(mean = mean(get(var), na.rm = T), sd = sd(get(var), na.rm = T))]
                } else {
                    zscore_reference = event[, list(mean = mean(get(var), na.rm = T), sd = sd(get(var), na.rm = T))]
                }
                
                pdff_reference = event[period == config$pdff_reference, list(mean = mean(get(var), na.rm = T), median = median(get(var), na.rm = T))]
                
                event[, (new.vars) := .(baseline[, mean], # baseline mean
                                        (get(var) - pdff_reference[, get(config$pdff_statistic)]) / pdff_reference[, get(config$pdff_statistic)] * 100, # %dF/F
                                        (get(var) - zscore_reference[, mean]) / zscore_reference[, sd], # z-score
                                        get(var) # mean
                                        )]
            }
            
            data$eventdata = rbindlist(list(data$eventdata, event))
        }
        dbg("processed event data contains %d total observations\n", data$eventdata[, .N])
        
        
    })
    
    # generate tables of normalized values and event averages
    observeEvent(data$eventdata, {
        df = copy(data$eventdata)
        data$pdff = data.table()
        for (var in input$data_columns) {
            data$pdff = rbindlist(list(data$pdff, dcast(df[, variable := var], variable + period ~ event, function(x){mean(x)}, value.var = paste0(var, "_pdff"))))
        }
        dbg("generated %%dF/F table\n")
    })
    
    observeEvent(data$eventdata, {
        data$zscore = data.table()
        df = copy(data$eventdata)
        for (var in input$data_columns) {
            data$zscore = rbindlist(list(data$zscore, dcast(df[, variable := var], variable + period ~ event, function(x){mean(x)}, value.var = paste0(var, "_zscore"))))
        }
        dbg("generated z Score table\n")
    })
    
    observeEvent(data$eventdata, {
        data$means = data.table()
        df = copy(data$eventdata)
        for (var in input$data_columns) {
            data$means = rbindlist(list(data$means, dcast(df[, variable := var], variable + period ~ event, function(x){mean(x)}, value.var = paste0(var, ""))))
        }
        dbg("generated means table\n")
    })
    
    observeEvent(data$eventdata, {
        data$avgs = data.table()
        df = copy(data$eventdata)
        df[, event_time := round(event_time, 6)] # rounding?
        for (var in input$data_columns) {
            for (var2 in c("", "_pdff", "_zscore")) {
                data$avgs = rbindlist(list(data$avgs, dcast(df[, variable := paste0(var, var2)], variable + event_time ~ event, value.var = paste0(var, var2))[, type := var2]))
            }
        }
        dbg("event averages\n")
    })
    
    # --------------------------------
    
    # render plot of full time series with some options
    output$overview_plot = renderPlot({
        plots = list()
        for (i in 1:length(input$data_columns)) {
            plot = ggplot(data$input, aes(x = time, y = get(input$data_columns[[i]]))) +
                labs(x = "Time (s)", y = input$data_columns[[i]]) +
                theme_bw() +
                theme(legend.position = "bottom", panel.grid = element_blank()) +
                scale_x_continuous(expand = c(0, 0), limits = c(input$overview_min, input$overview_max)) +
                scale_y_continuous(expand = c(0, 0)) 
            
            if ("Event Markers" %in% input$overview_options) {
                plot = plot + geom_vline(data = data$events, mapping = aes(xintercept = time), linetype = "dotted", na.rm = T)
            }
            
            if ("Lines" %in% input$overview_options) {
                plot = plot + geom_line(size = 0.5, color = "#c0c0c0", na.rm = T)
            }
            
            if ("Points" %in% input$overview_options) {
                plot = plot + geom_point(size = 0.5, na.rm = T, alpha = 0.5)
            }
            plots[[i]] = ggplotGrob(plot)
        }
        grid.arrange(grobs = plots, ncol = 1)
    }, height = function() {
        session$clientData$output_overview_plot_width * 10 / 16 * length(input$data_columns)
    })
    
    # render density plot showing activity across all events centered on t=0
    output$density_plot = renderPlot({
        plots = list()
        for (i in 1:length(input$data_columns)) {
            plot = ggplot(data$eventdata, aes(x = event_time, y = get(input$data_columns[[i]]))) + 
                geom_bin2d(bins = config$plot_density_bins, na.rm = T) +
                geom_vline(aes(xintercept = 0), linetype = "dotted") +
                labs(x = "Event Time (s)", y = input$data_columns[[i]]) +
                theme_bw() +
                theme(legend.position = "bottom") +
                scale_x_continuous(expand = c(0, 0)) +
                scale_y_continuous(expand = c(0, 0)) + 
                scale_fill_gradientn(colors = config$plot_gradient, guide = guide_colorbar(barwidth = 14, barheight = 0.75))
            plots[[i]] = ggplotGrob(plot)
            dbg("generated density plot for %s\n", input$data_columns[[i]])
        }
        grid.arrange(grobs = plots, ncol = 1)
    }, height = function() {
        session$clientData$output_density_plot_width * 10 / 16 * length(input$data_columns)
    })
    
    output$events_table = renderDT({
        data$events[, .(time, event)]
    }, selection = "single", options = list(lengthChange = F, pageLength = 5, searching = F, pagingType = "numbers"), rownames = F, style = "bootstrap", class = "table-bordered table-condensed")
    
    # render selected event with shading for periods
    output$events_plot = renderPlot({
        req(input$events_table_rows_selected)
        plots = list()
        for (i in 1:length(input$data_columns)) {
            
            plot = ggplot(data$eventdata[event == data$events[input$events_table_rows_selected, event],], aes(x = event_time, y = get(input$data_columns[[i]]))) + 
                geom_vline(mapping = aes(xintercept = 0), linetype = "dotted") + 
                labs(x = "Event Time (s)", y = input$data_columns[[i]]) + 
                theme_bw() + 
                theme(legend.position = "bottom") +
                scale_x_continuous(expand = c(0, 0)) + scale_y_continuous(expand = c(0, 0))
            
            
            if ("Shade Intervals" %in% input$events_options) {
                plot = plot + geom_rect(aes(xmin = start, xmax = end, ymin = -Inf, ymax = Inf, fill = factor(name)), config$periods[config$periods[, "name"] != "unassigned",], alpha = 0.5, inherit.aes = F) +
                    scale_fill_brewer(palette = "Pastel1", guide = guide_legend(title = "Period", direction = "horizontal", nrow = 1, override.aes = list(size = 4)))
            }
            
            if ("Lines" %in% input$events_options) {
                plot = plot + geom_line(aes(x = event_time, y = get(input$data_columns[[i]])), inherit.aes = F, size = 0.5, color = "#c0c0c0", na.rm = T)
            }
            
            if ("Points" %in% input$events_options) {
                plot = plot + geom_point(aes(x = event_time, y = get(input$data_columns[[i]])), inherit.aes = F, alpha = 0.25, color = "#000000", size = 0.5, na.rm = T)
            }
            
            plots[[i]] = ggplotGrob(plot)
        }
        grid.arrange(grobs = plots, ncol = 1)
    }, height = function() {
        session$clientData$output_events_plot_width * 10 / 16 * length(input$data_columns)
    })
    
    # boxplot of normalized values and stuff
    output$summary_boxplot = renderPlot({
        dat = NULL
        if (input$summary_type == "Raw") {
            dat = melt(data$means[period != "unassigned",], id = c("variable", "period"), variable.name = "event")
        } else if (input$summary_type == "%dF/F") {
            dat = melt(data$pdff[period != "unassigned",], id = c("variable", "period"), variable.name = "event")
        } else if (input$summary_type == "z Score") {
            dat = melt(data$zscore[period != "unassigned",], id = c("variable", "period"), variable.name = "event")
        }
        if (!is.null(dat)) {
            intervals = as.character(config$periods[, "name"])
            intervals = intervals[intervals != "unassigned"]
            dat[, period := ordered(period, levels = intervals)]
            
            dodge = position_jitterdodge()
            plot = ggplot(dat, aes(x = period, y = value, fill = variable, label = event)) +
                theme_bw() +
                stat_summary(fun.data = MinMeanSEMMax, geom = "boxplot", na.rm = T, position = position_dodge()) +
                scale_fill_discrete(guide = guide_legend(title = "Variable")) +
                labs(y = input$summary_type)
            
            if ("Labels" %in% input$summary_options) {
                plot = plot + geom_text_repel(min.segment.length = 0, box.padding = 0.4, na.rm = T, position = position_jitterdodge(seed = 1, jitter.width = 0.25), segment.alpha = 0.333)
            }
            if ("Points" %in% input$summary_options) {
                plot = plot + geom_point(na.rm = T, position = position_jitterdodge(seed = 1, jitter.width = 0.2))
            }
            plot
        }
    })
    
    # summary tables to display
    output$summary_table = renderRHandsontable({
        dat = NULL
        if (input$summary_type == "Raw") {
            dat = data$means
        } else if (input$summary_type == "%dF/F") {
            dat = data$pdff
        } else if (input$summary_type == "z Score") {
            dat = data$zscore
        }
        
        if (!is.null(dat)) {
            hotdata = copy(dat)[period != "unassigned",][, `:=`(Average = rowMeans(.SD, na.rm = T)), by = .(variable, period)][, `:=`(`Average (sans #1)` = rowMeans(.SD, na.rm = T)), by = .(variable, period), .SDcols = -c("1")]
            hot = rhandsontable(hotdata, readOnly = T)
            hot = hot_col(hot, 3:(dim(dat)[2]+2), format = "0.6")
            hot = hot_cols(hot, fixedColumnsLeft = 2)
            hot
        }
    })
    
    output$period_table = renderRHandsontable({
        hot = rhandsontable(config$periods, rowHeaders = NULL) %>%
            hot_col("name", type = "autocomplete") %>%
            hot_col(c("start", "end"), format = "0.5")
        hot
    })
    
    # save edited periods as csv file
    output$period_download = downloadHandler(
        filename = "periods configuration.csv",
        content = function(file) {
            write.csv(config$periods, file, row.names = F)
        }
    )
    
    # someone clicked download, build excel file as output
    output$output_download = downloadHandler(
        filename = function() {
            paste0(strsplit(input$input_file$name, ".", fixed = T)[[1]][1], " (Processed).xlsx")
        },
        
        content = function(file) {
            withProgress(message = "Creating file", value = 0, detail = "Aggregating data...", {
                sheets = list(`Event Data` = data$eventdata)
                
                incProgress(0.1)
                
                # events by time by normalization
                df = data$avgs[, Average := rowMeans(.SD, na.rm = T), by = .(type, variable, event_time)][, `Average (sans #1)` := rowMeans(.SD, na.rm = T), by = .(type, variable, event_time), .SDcols = -c("1")]
                sheets[["Events by Time (Raw)"]] = df[type == "",]
                sheets[["Events by Time (dFF)"]] = df[type == "_pdff",]
                sheets[["Events by Time (zScore)"]] = df[type == "_zscore",]
                
                incProgress(0.1)
                
                # average event values table
                sheets[["Raw by Event"]] = data$means[, `:=`(Average = rowMeans(.SD, na.rm = T)), by = .(variable, period)][, `:=`(`Average (sans #1)` = rowMeans(.SD, na.rm = T)), by = .(variable, period), .SDcols = -c("1")][period != "unassigned",]
                sheets[["dFF by Event"]] = data$pdff[, Average := rowMeans(.SD, na.rm = T), by = .(variable, period)][, `:=`(`Average (sans #1)` = rowMeans(.SD, na.rm = T)), by = .(variable, period), .SDcols = -c("1")][period != "unassigned",]
                sheets[["zScore by Event"]] = data$zscore[, Average := rowMeans(.SD, na.rm = T), by = .(variable, period)][, `:=`(`Average (sans #1)` = rowMeans(.SD, na.rm = T)), by = .(variable, period), .SDcols = -c("1")][period != "unassigned",]

                incProgress(0.1)
                
                # put a blank line in the output between events and variables for readability
                for (p in names(sheets)) {
                    new.dt = sheets[[p]][.0]
                    blank.line = rbindlist(list(new.dt, as.list(rep(NA, length(names(sheets[[p]]))))))
                    if (p == "Event Data") {
                        for (e in sheets[[p]][, unique(event)]) {
                            new.dt = rbindlist(list(new.dt, sheets[[p]][event == e,], blank.line))
                        }
                    } else {
                        for (v in sheets[[p]][, unique(variable)]) {
                            new.dt = rbindlist(list(new.dt, sheets[[p]][variable == v,], blank.line))
                        }
                    }
                    sheets[[p]] = new.dt
                }
                
                incProgress(0.35, detail = "Saving...")
                
                # saved!
                write_xlsx(sheets, file, col_names = T)
                
                incProgress(0.35, detail = "Done!")
                
                dbg("wrote output file\n")
            })
        }
    )
}

shinyApp(ui = ui, server = server)