cat:
	cat Makefile

split:
	cargo run -- split module.bas

concat:
	cargo run -- concat
	
cmt:
	git add .
	git commit --allow-empty-message -am ""

commit: cmt concat
	make cmt
