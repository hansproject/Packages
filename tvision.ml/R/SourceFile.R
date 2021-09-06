# Hello, world!
#
# This is an example function named 'hello'
# which prints 'Hello, world!'.
#
# You can learn more about package authoring with RStudio at:
#
#   http://r-pkgs.had.co.nz/
#
# Some useful keyboard shortcuts for package authoring:
#
#   Install Package:           'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

#' @importFrom dplyr %>%
#' @export
hello <- function() {
  print("Hello, world!")
}

# read_file function which reads only csv files
# Accepts 2 arguments:
# @ path     : path where the file is located
# @ file_name: name of file to be read
#' @export
#------------------------------------------------------
getDBConnection= function(path){
  return(DBI::dbConnect(RSQLite::SQLite(), "C:/Users/ndeff/OneDrive/Desktop/Research and online classes/Company Projects/TVision/db.sql3"))
}

# pre_process function which preprocess the current dataset and accept 1 argument. Then it returns the pre-processed data
# @ data       : Data to be preprocessed
#' @export
#-------------------------------------------------------------------
pre_process= function(data, cols){

  for(i in cols){
    data[,i]<- as.numeric(as.factor(unlist(data[,i])))
  }

  return(data)
}

# initialize_spark function
# Accepts 2 arguments
# - n_cores     : number of cores to be used by Spark
# - memory_taken: amount of memory to be taken from your RAM. e.g: 4G will mean 4 Go
#' @export
#-------------------------------------------------------------------------------------
initialize_spark= function(n_cores, memory_taken){
  conf <- sparklyr::spark_config()
  conf$`sparklyr.cores.local` <- n_cores
  conf$`sparklyr.shell.driver-memory` <- memory_taken

  # Then create and connect a Spark cluster that will have the config above with this command
  system.time( sc <- sparklyr::spark_connect(master = "local",version = "2.2.0",
                                   config = conf))
  return (sc)
}

# train function
# Accepts 4 arguments
# - model_name : which model do you want to use for training
# - source_file: Is it location or pbp dataset being used
# - formula    : Input formula to train the model
# - train_data : data used for training
#' @export
#-------------------------------------------------------------------
train= function(model_name, destination_path, formula, train_data){

  if(strcmp(model_name, "Random Forest")){

    rf_spark<- sparklyr::ml_random_forest_regressor(train_data, formula, num_trees = 100)
    saveRDS(rf_spark,file= paste(destination_path,"rf_spark",".rds",sep = ""))
    return(rf_spark)

  }

  if(strcmp(model_name, "Linear")){

    lr_spark<- sparklyr::ml_linear_regression(train_data, formula)
    saveRDS(lr_spark,file= paste(destination_path,"lm_spark",".rds",sep = ""))
    return(lr_spark)

  }

  if(strcmp(model_name, "Perceptron")){

    perceptron_spark<- sparklyr::ml_multilayer_perceptron_classifier(train_data,  ft_r_formula(train_data, formula, label_col = "attention_to_duration_ratio"), layers= c(23,9,1), max_iter=50)
    saveRDS(perceptron_spark,file= paste(destination_path,"perceptron_spark",".rds",sep = ""))
    return(perceptron_spark)

  }

  else
    return(NULL)
}

# show_features_importance function which display features importances graph comparaison
# Accepts 2 arguments:
# - models     : A list of trained models
# - graph_title: Title of the graph
#' @importFrom dplyr %>%
#' @export
#----------------------------------------------------------
show_features_importance <- function(models, graph_title){
  ml_models <- models

  # Extract features with their respectives coefficients
  feature_lm <-  data.frame(feature= ml_models["Linear"]$Linear$feature_names, importance= abs(ml_models["Linear"]$Linear$coefficients[2:length(ml_models["Linear"]$Linear$coefficients)]), row.names = NULL)

  feature_lr <-  data.frame(feature= names(ml_models["Lasso"]$Lasso$beta[,position]), importance= abs(ml_models["Lasso"]$Lasso$beta[,position]), row.names = NULL)

  feature_importance <- data.frame()

  # Calculate feature importance
  for(i in c( "Linear","Lasso","Random Forest")){

    if(strcmp(i,"Linear")){
      feature_importance <- feature_lm %>%
        mutate(Model = i) %>%
        rbind(feature_importance, .)

    }
    else if(strcmp(i,"Lasso")){
      feature_importance <- feature_lr %>%
        mutate(Model = i) %>%
        rbind(feature_importance, .)
    }
    else{
      feature_importance <- sparklyr::ml_tree_feature_importance(ml_models[[i]]) %>%
        mutate(Model = i) %>%
        rbind(feature_importance, .)
    }
  }

  # Plot results
  feature_importance %>%
    ggplot2::ggplot(aes(reorder(feature, importance), importance, fill = Model)) +
    ggplot2::theme(plot.title = element_text(size = 15, hjust = 0.5), panel.background = element_rect("white"))+
    ggplot2::facet_wrap(~Model) +
    ggplot2::geom_bar(stat = "identity") +
    ggplot2::coord_flip() +
    ggplot2::xlab("") +
    ggplot2::ggtitle(graph_title)
}
