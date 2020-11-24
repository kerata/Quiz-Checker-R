Sys.setenv(LANG = "en")

suppressWarnings(suppressMessages(library(R.utils)))
MULT_ERR <- "MULTIPLE_ERROR"

c_equivalence <- list("numeric", "integer")
c_conversion <- c("as.integer", "as.numeric")
c_exception <- c(NA, function(rhs, lhs) abs(rhs - lhs) < 1e-5)

# EXPERIMENTAL
class_expansion <- function(item) {
  # Automatically expand equivalent class results

  # Find location of the class
  loc <- match(class(item), c_equivalence, nomatch = -1)
  if(loc == -1) return(item) # Return as is if does not have equivalent class

  counterpart <- do.call(c_conversion[loc], list(item))
  if(!is.na(c_exception[loc]) &&
     do.call(c_exception[loc][[1]], list(item, counterpart))) {
    # Check exception part
    # Only whole ones Ex: 1.5 != 1L, 1.0 == 1L
    # Class equivalence
    return(item)
  }
  else return(list(item, counterpart))
  # Or expand the class with new equivelant class
}

# EXPERIMENTAL
expand <- function(val, base) {
  # Expand the output list with equivelant class

  # For naming convention
  counter <- 1

  res <- list()
  for (i in seq_along(val)) {
    # Expand the item if possible
    add <- class_expansion(val[[i]])
    # If item stays same return as is when val is not a list
    if(identical(add, val[[i]]) && !is.list(val)) {
      return(val)
    }

    # Otherwise, calculate its name and append to resulting list
    names(add) <- paste(rep(base, length(add)),
                        counter:(counter+length(add)-1), sep="")
    counter <- counter+length(add)
    res <- c(res, add)
  }
  return(res)
}

do_test <- function(target_name, value, output_answer) {
  # check whether outputs are identical and record
  # Control if there are multiple valid answers if value is list and sub elements are numerated starting with variable names
  # Ex: <value name>: "series", <list element names>: "series1", "series2"

  # Since `expand` is compatible with naming convention
  # expanding the answer beforehand should work!
  # EXPERIMENTAL
  # value <- expand(value, target_name)

  if(is.list(value) && length(names(value)) > 0 &&
     all(sapply(names(value), startsWith, target_name)) &&
     all(sort(as.integer(sapply(names(value), substring, nchar(target_name)+1))) == 1:length(value))) {
    return(any(sapply(value, identical, output_answer)))
  }
  else {
    return(identical(output_answer, value))
  }
}

get_output <- function(is_fnc, is_mult_in, target_name, fnc_input, m_envir) {
  # Get output from answer, function or variable
  if(is_fnc) {
    if(is_mult_in) {
      output_answer <- list()
      for(i in 1:length(fnc_input)) {
        # When we try functions for more than one times, there is a chance for implemented
        # function to fail on certain inputs and not the other.
        # If function crashes for all trials, then throw exception...
        tryCatch({
          expr <- expression(dump <- capture.output(outt <- do.call(target_name, fnc_input[[i]], envir = m_envir)))
          withTimeout(expr, substitute = F, timeout = 1, onTimeout = "error")
          output_answer[[i]] <<- outt
        }, error = function(e) {
          output_answer[[i]] <<- MULT_ERR
          if(length(fnc_input) == length(output_answer) &&
             all(sapply(output_answer, identical, MULT_ERR))) {
            e$message <- paste(paste("get_output", MULT_ERR, sep="|"), e$message, sep="::")
            stop(e)
          }
        })
      }
    } else {
      expr <- expression(dump <- capture.output(output_answer <- do.call(target_name, fnc_input, envir = m_envir)))
      withTimeout(expr, substitute = F, timeout = 1, onTimeout = "error")
    }
  } else {
    output_answer <- get(target_name, envir = m_envir)
  }
  return(output_answer)
}

is_script_compatible <- function(path, unsupported="readline") {
  suppressWarnings(content <- readLines(path))
  if(length(grep(unsupported, content)) != 0) {
    stop(paste("UnsupportedFunctionUsage : ", unsupported))
  }
  if(length(content) == 1 && content=="# AUTO GENERATED ANSWER FOR NULL SUBMISSION"){
    stop("NULL SUBMISSION")
  }
}

initialize_answers_df <- function(answer_list, student_ids, count_of_answers, count_of_runs) {
  # create an empty object to collect results and error messages if any for profiling purposes
  # this is a global object in order to collect error messages
  # first column shows whether the output of the code is identical to the output provided
  # second column collects any error messages
  answers_df <<- data.frame(bool = rep(NA, count_of_answers),
                            error = character(count_of_answers), stringsAsFactors = F)

  # answers_df$bool <<- NA # replace the first column with NA's since it is initiated with FALSE values

  # This is taken as strings to account for any non-numeric characters

  # replicate the df as much as the runs + 1. + 1 is for final result and source time error columns (for multiple runs, all runs must yield T)
  answers_df <<- as.data.frame(rep(answers_df, count_of_runs + 1), stringsAsFactors = F)

  # and add student id column
  answers_df <<- cbind(student_ids = student_ids, answers_df, stringsAsFactors = F)

  names(answers_df)[2:3] <<- c("final_result", "source_errors")
  names(answers_df)[seq(4, length(answers_df), 2)] <<- names(answer_list)

  answers_df <<- cbind(answers_df, points=0, stringsAsFactors = F)

  return(answers_df)
}

load_target_script <- function(file_path, loader=NULL) {
  envirs <- list()
  tryCatch({
    is_script_compatible(file_path)
    enforce_inj <- FALSE
    if(!is.null(loader)) {
      student_inj_env <- attach(what=NULL, name="student_inj_env")
      dump <- capture.output(enforce_inj <- loader(file_path, student_inj_env))
      envirs[["student_inj_env"]] <- student_inj_env
    }
    if(is.logical(enforce_inj) && !enforce_inj) {
      student_env <- attach(what=NULL, name="student_env")
      # dump <- capture.output(sys.source(file_path, student_env)) # source the input file
      # source the input file, capture the output so it will not be echoed to any callers
      # use loader function if provided for given tests
      withTimeout(expression(dump <- capture.output(sys.source(file_path, student_env))),
                  substitute = F, timeout = 1.5, onTimeout = "error")

      envirs[["student_env"]] <- student_env
    }
  }, error = function(e) {
    if(exists("student_env")) {
      detach("student_env")
    }
    if(exists("student_inj_env")) {
      detach("student_inj_env")
    }

    if(is(e,"TimeoutException")) {
      stop(paste("load_target_script", e$message, sep="::"))
    } else {
      e$message <- paste("load_target_script", e$message, sep="::")
    }
    stop(e)
  })

  # Delete the created plots if exists...
  graphics.off()
  return(envirs)
}

evaluate_solution <- function(answer_list, run) {
  calculated <- list()

  # These are calculated for the answer_df, which is designed for class control
  calculated$bool_ind <- 2 + run*2 # column index for the boolean value of current run
  calculated$error_ind <- calculated$bool_ind + 1  # column index for the error message of current run

  calculated$target_name <- names(answer_list)[run] # the input for current run
  calculated$is_inj <- startsWith(calculated$target_name, "INJ__")
  if(calculated$is_inj) calculated$target_name <- substring(calculated$target_name, 6)

  calculated$is_fnc <- startsWith(calculated$target_name, "FNC__")

  calculated$is_mult_in <- F

  if(calculated$is_fnc) {
    calculated$target_name <- substring(calculated$target_name, 6)

    if(length(answer_list[[run]]) == 2) {
      calculated$fnc_input <- answer_list[[run]][[1]]
      calculated$value <- answer_list[[run]][[2]]
    } else {
      calculated$fnc_input <- answer_list[[run]][c(T, F)]
      calculated$value <- answer_list[[run]][c(F, T)]

      calculated$is_mult_in <- T
    }
  } else {
    calculated$value <- answer_list[[run]] # the output for current run
  }

  return(calculated)
}

evaluate_student_answer <- function(file_path, answer_list, loader=NULL) {
  # Load given answer script for evaluation
  with(load_target_script(file_path, loader), {
    t_names <- NULL
    indices <- NULL
    values <- NULL
    for (run in 1:length(answer_list)) {
      # Parse given solution for evaluation
      with(evaluate_solution(answer_list, run), {
        # Update result content
        indices <<- c(indices, bool_ind)
        t_names <<- c(t_names, target_name)
        tryCatch({
            # get the variable in the answer for the current round
            if(is_inj) {
              if(typeof(get(target_name, envir = student_inj_env)) != "closure" && is_fnc)
                throw(Exception(sprintf("%s has to be a function!", target_name)))
              output_answer <- get_output(is_fnc, is_mult_in, target_name, fnc_input, student_inj_env)
            } else {
              if(typeof(get(target_name, envir = student_env)) != "closure" && is_fnc)
                throw(Exception(sprintf("%s has to be a function!", target_name)))
              output_answer <- get_output(is_fnc, is_mult_in, target_name, fnc_input, student_env)
            }

            if(is_mult_in) {
              res <- c()
              for(i in 1:length(value)) {
                res <- c(res, do_test(target_name, value[[i]], output_answer[[i]]))
              }

              values <<- c(values, any(res))
            } else if(typeof(value) == "closure") {
              # Use given funtion to override the testing
              values <<- c(values, value(fnc_input, output_answer))
            } else {
              values <<- c(values, do_test(target_name, value, output_answer))
            }
          }, error = function(e) {
            values <<- c(values, F)
            indices <<- c(indices, error_ind)
            values <<- c(values, as.character(e$message))
          }
        )
      })
    }

    if(exists("student_env")) {
      detach("student_env")
    }
    if(exists("student_inj_env")) {
      detach("student_inj_env")
    }

    return(list(indices=indices, values=values, t_names=t_names))
  })
}

## Adapted from first version, designed to test variable values
check_quizzes <- function(name = "week01", section = "01", question = "01", answers_path = NULL,
                          data_path = "io", results_path = "results",
                          answer_list = NULL, max_point = 20, should_output_file = T, recalculate = F)
{
  # construct output file names, "out" name in the beginning and RData in the
  # end is important. As well as csv outputs since git is designed to ignore
  # these outputs automatically. Please be mindful before changing these parts.
  data_file <- sprintf("%s/out_%s%s_q%s.RData", data_path, name, if(is.null(section)) "" else sprintf("_sect%s", section), question)
  if (file.exists(data_file) && !recalculate) {
    load(data_file)
    return(answers_df)
  }

  # create answer path from name, section and question if nothing specific has given
  if (is.null(answers_path))
  {
    answers_path <- sprintf("answers/%s%s/q%s", name, if(is.null(section)) "" else sprintf("/sect%s", section), question)
  }

  results_file <- sprintf("%s/%s%s_q%s.csv", results_path, name, if(is.null(section)) "" else sprintf("_sect%s", section), question)

  if (is.null(answer_list)) # if the input list is null
  {
    # load the input list from RData file and assign into answer_list
    answer_list_name <- sprintf("in_%s%s_q%s", name, if(is.null(section)) "" else sprintf("_sect%s", section), question)
    answer_list_filename <- sprintf("%s/%s.RData", data_path, answer_list_name)
    do.call(load, list(answer_list_filename))
    assign("answer_list", get(answer_list_name))
    rm(answer_list_name)
  } else if (!is.list(answer_list)) {
    if (file.exists(answer_list)) {
      load(answer_list)
      answer_list_name <- sprintf("in_%s%s_q%s", name, if(is.null(section)) "" else sprintf("_sect%s", section), question)
      assign("answer_list", get(answer_list_name))
      rm(list = c(answer_list_name, "answer_list_name"))
    }
  }

  if(exists("loader")) {
    assign("loader", loader[[1]])
  } else {
    loader <- NULL
  }

  files <- list.files(answers_path, "*.R") # get list of answer files
  student_ids <- gsub("(?i).R$", "", files) # clear the extension part of the filenames so only student id's remain.
  count_of_answers <- length(files) # number of answers. note that no object other than answers should be kept under the input path

  count_of_runs <- length(answer_list) # number of runs from input

  # Makes initialization in callers environment, therefore answers_df exists in this scope as well
  # and it is the one to be updated...
  initialize_answers_df(answer_list, student_ids, count_of_answers, count_of_runs)

  # for1, outer loop, across answer files
  for (ans in 1:count_of_answers) {
    answer_file_path <- sprintf("%s/%s", answers_path, files[ans]) # create filename to be sourced

    tryCatch({
      with(evaluate_student_answer(answer_file_path, answer_list, loader), {
        answers_df[ans, indices] <<- values
      })
    }, error = function(e) {
      error_group <- strsplit(e$message, "::", useBytes = T)[[1]]
      wrapper_func <- error_group[1]
      error_msg <- error_group[2]
      caller_func <- e$call[[1]]

      if(wrapper_func == "load_target_script") {
        answers_df[ans, "final_result"] <<- F # the final result if F
        answers_df[ans,"source_errors"] <<- as.character(error_msg)
      } else {
        e$message <- paste("Unknown", error_msg, sep="::")
        stop(e)
      }
    })
  } # close for1

  # collect the results of all runs to final result
  # NA's converted to FALSE
  target_rows <- is.na(answers_df$final_result)
  results2calc <- answers_df[target_rows, 2 + 2*(1:count_of_runs)]
  if(is.vector(results2calc)) {
    answers_df[target_rows,"final_result"] <<- results2calc
  } else {
    answers_df[target_rows,"final_result"] <<- apply(results2calc, 1,
                                                     function(x)
                                                     {
                                                       if (is.na(all(x)))
                                                       {
                                                         return(F)
                                                       }
                                                       else
                                                       {
                                                         return(all(x))
                                                       }
                                                     }
    )
  }

  # Calculate total grade
  results2calc <- answers_df[,2 + 2*(1:count_of_runs)]
  if(is.vector(results2calc)) {
    correct_ones <- sapply(results2calc, function(a) { if(is.na(a) || a == F) F else T })
    answers_df[,"points"] <<- correct_ones*max_point/count_of_runs
  } else {
    answers_df[,"points"] <<- apply(results2calc, 1,
                                    function(x)
                                    {
                                      y <- sapply(x, function(a) { if(is.na(a) || a == F) F else T })
                                      return(sum(y)*max_point/count_of_runs)
                                    }
    )
  }

  # error columns for which new line feeds will be replaced by whitespaces or quotes will be deleted
  error_cols <- 1 + 2 * (1:(count_of_runs + 1))

  # replace new lines with whitespaces, delete double or single quotes inside error messages in order to better parse with excel, column, awk, etc
  for (ercols in error_cols)
  {
    textt <- gsub("\n", " ", answers_df[, ercols])
    textt <- gsub("\"", "", textt)
    textt <- gsub("\'", "", textt)
    answers_df[, ercols] <<- textt
  }

  if(should_output_file) {
    # save the data as RData.
    save(answers_df, file = data_file)
    # write into csv file
    write.csv2(answers_df[order(answers_df[,1], decreasing = T),], file = results_file, row.names = F)
  }

  return(answers_df)
}

# Build arguments
args <- commandArgs(trailingOnly=TRUE)
if(length(args) != 0) {
  quiz_args <- args[c(F, T)]
  names(quiz_args) <- args[c(T, F)]

  quiz_args <- as.list(quiz_args)

  pos <- match("max_point", names(quiz_args))
  quiz_args[[pos]] <- as.integer(quiz_args[pos])

  for (logic in c("TRUE","FALSE","T","F")) {
    pos <- match(logic, quiz_args)
    if(!is.na(pos)) quiz_args[[pos]] <- as.logical(quiz_args[[pos]])
  }

  pos <- match("section", names(quiz_args))
  if(quiz_args[pos] == "NULL") quiz_args[pos] <- list(NULL)

  # Call the function and return resulting dataframe to caller...
  print.data.frame(do.call("check_quizzes", quiz_args), quote = T)
} else {
  rm(args)
}
