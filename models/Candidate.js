const mongoose = require("mongoose");

const CandidateSchema = new mongoose.Schema({
  name: String,
  phone: String,
  method: String,
  language: String,

department: {
    type: String,
    default: "Not selected"
  },

  subDepartment: {
    type: String,
    default: "Not selected"
  },

  section: {
    type: String,
    default: "Not selected"
  },

  designation: {
    type: String,
    default: "Not selected"
  },

  status: {
    type: String,
    default: "Pending"
  },

  interviewDate: String,
  interviewTime: String,

  score: {
    type: Number,
    default: 0
  },

  totalQuestions: {
    type: Number,
    default: 0
  }
}, {
  timestamps: true
});

module.exports = mongoose.model("Candidate", CandidateSchema);