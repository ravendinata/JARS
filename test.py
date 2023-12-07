import processor.grader_report as gr
import processor.semester_report as sr

# source_6 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 6 Items.xlsm")
# source_7 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 7 Items.xlsm")
# source_9 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 9 Items.xlsm")
source_10 = gr.GraderReport("Z:\Work\JAC\Project JARS\Playground GR - 10 Items.xlsm")

# sr.Generator("Z:\Work\JAC\Project JARS\PG Test\Six", source_6).generate_all(autocorrect = False, force = True)
# sr.Generator("Z:\Work\JAC\Project JARS\PG Test\Seven", source_7).generate_all(autocorrect = True, force = True)
# sr.Generator("Z:\Work\JAC\Project JARS\PG Test\\Nine", source_9).generate_all(autocorrect = False, force = True)
sr.Generator("Z:\Work\JAC\Project JARS\PG Test\Ten", source_10).generate_all(autocorrect = False, force = True)