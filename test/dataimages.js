'use strict';

var battery1 = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAYAAACinX6EAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABq0lEQVRoge2YsS4EYRDHf3PuktNcwu2KwlVEofMCJCQKnSfwAPS8Alp6b0BDVCIqKoXoSMiJhPtOVJycu1HcilC4zO5e1rK/5tsvmf/M5J9vsslARkbGf0bsEs0x+DiL6Hg4fYyIvtLXOuF++Dx0Clu45vFre6jMhS3YAxSVVereehixzYCyW0R0O7jdAG9hisZICfCBNlDB+XfWBHlj/GRwHuP8aWux2Bmp9tMoPgM5VCYAswE5W3S7Y5jKi7VQT7itfPbx0ZuRUKIveLXLrjFv+VmeBm4i1+oB0Q2A0a4RhWYhhjo/o7JK2RWpe7sWmW0EfjfTiO5QdjMWUfoNEN1CdAuoBvd5izyOEUiW2tAyAGU3hmgF0aJFnv4XEJHMgKQbSJrMgKQbSJr0/wX8h00AVCc6pzQs8vQboLL07b5vkf+lEThCZYG6d2gRxfECrrpGNAvNGOr8jOgazj+wyqIb4PyxyDkSxDYC7VxnAyTa34tmzIxUP/v46M2I9QWcBecUXu2apFdiDUrBVxvRizAp/sZSFFZw/kYYcdrX4g36WqdR1uIZGRn/m3fCBW+4FEtAXQAAAABJRU5ErkJggg==';
var battery2 = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAYAAACinX6EAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABmUlEQVRoge2YsUoDQRCGv0ksYpPC3ImFqQwW6XwBBQULO5/AB9BeX0FtTe8baKNYiVhpZSF2CkpE0GwsNYJmLHIhaDe3F84z9zWbhf1nfn52OTKQk5MzyohdogUmXpcQnY2nTxDRD4pfFzxPXccuYTuuY4StI1SW4zYcAorKFu1gJ47YFkDFrSG6H+0egM84TROkDIRAF6jiwidrgTHj+bloPceFC9ZmiTPdHKdTegMKqNQBcwAF2+luLzCVd2ujofBYHfjoezMSS/SDoHXrpXdhzduDB/4BwEwCNfxR2aLiSrSDQ4vM9gT+NguIHlBxixZR9gMQbSDaAJrRfsUiT+IJpEtrcgOAiqshWkW0ZJFn/wZ4kgeQtoG0yQNI20DaZP8rEL7sAaBa763SscizH4DK+q/9sUX+n57AGSqrtINTiyiJG3CXQA1/RLdx4YlV5h9Ayv/mfLE9gW6hNwESHR+GGTPTzYGPvjcj1htwFa3zBK170h6JdShHv7qI3sQp8T+GorCJC3fjiLM+Fu9Q/Lr0GYvn5OSMNt8TO2q4lNWMWQAAAABJRU5ErkJggg==';
var battery3 = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAYAAACinX6EAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABoUlEQVRoge2YsS4EURSGv7MUq1HYuaKwFVHovAAJiULnCTwAvX0FtPTegIaoRFRUCtGRsCsS9q6SlbBHsbOEbHNmZjNmzdfcucn5zzn5c24mOZCTk/OfEbtEC4w8LyA6FU2fIKJvDHyc8Th2GTmFLVwHcfUDVBajFuwBikqFRrAZRWwzoORXEN0Nb3fAe5SiCTIMOKAFlPHuwZpg0Bg/E56neDdnLZY447UhmsUXoIDKNGA2oGCLbrUNU3m1FuoJ9+XvPjq9GYkk+kFQv46doztVvJvvUe4v4hsAEwnkiI9KhZIv0gj2LTLbE/jbzCG6R8mbpib7BojuILoD1ML7kkWexBNIl/roGgAlP4loGdGiRZ79CYhJbkDaDaRNbkDaDaRN9v8C7mkbANXp9ilNizz7Bqis/rofWuT99AROUFmmERxbRElMwE0CObpRNUWLbuDdkbVIfAO8m4ydI0VsT6BVaG+ARId60YyZ8dp3H53ejFgn4CI8Zwnqt6S9EmsyHH61EL2KkqI/lqKwjndbUcRZX4s3Gfg4j7MWz8nJ+d98AsKVbJq58zg+AAAAAElFTkSuQmCC';

module.exports = function (excel) {
	var workbook = excel.createWorkbook();

	workbook.addImage(battery1, 'png', 'battery1');
	workbook.addWorksheet()
		.setData([
			[2541, 'Nullam aliquet mi et nunc tempus rutrum.', 260, {image: 'battery1'}],
			[2542, 'Nullam aliquet mi et nunc tempus rutrum.', 260, {image: {data: battery2, type: 'png'}}],
			[2543, 'Nullam aliquet mi et nunc tempus rutrum.', 260, {image: {image: 'battery1', config: {left: 5}}}],
			[2544, 'Nullam aliquet mi et nunc tempus rutrum.', 260,	{image: {image: {data: battery3, type: 'png'}}}],
			[2545, 'Nullam aliquet mi et nunc tempus rutrum.', 260,
				{
					value: 777,
					image: {image: {data: battery3, type: 'png'}, config: {cols: 2, rows: 3, left: 5, top: 5, right: 5, bottom: 5}}
				},
				42]
		]);

	return workbook;
};
