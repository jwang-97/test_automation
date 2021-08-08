```flow
st=>start: Case result
e=>end: send email(attach the report file)
op1=>operation: main.py
op2=>operation: Generate report ppt
sub1=>subroutine: export Contact & Autoforce by python
op3=>operation: save_comment|current
op4=>operation: set_sentiment|current
op5=>operation: set_record|current

cond1=>condition: last type of cases?
cond2=>condition: proxy_list空?
cond3=>condition: ids_got空?
cond4=>condition: 爬取成功??
cond5=>condition: ids_remain空?

io1=>inputoutput: call SAM Macro to export Static plots
io2=>inputoutput: call Multi Macro to export Dynamic plots
io3=>inputoutput: ids-got

st->op1(right)->io1->cond1
cond1(yes)->sub1->io2->op2
cond2(no)->io1
cond2(yes)->sub1
cond1(no)->io1
cond4(yes)->io3->cond3
cond4(no)->io1
cond3(no)->op4
cond3(yes, right)->cond5
cond5(yes)->op5
cond5(no)->cond3
op2->e

```