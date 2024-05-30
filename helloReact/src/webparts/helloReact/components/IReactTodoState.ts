import {ITodoItem } from "../ITodoItem";

export interface IReactTodoState {
    todoItems?: ITodoItem[];
    showNewToPanel?: boolean;
    newItemTitle?: string;
    newITemDone?: boolean;
}